"""
Discord Sentiment Analysis Dashboard
Loads messages CSV, runs Claude sentiment analysis,
and displays an interactive Streamlit dashboard.
"""

import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from claude_sentiment import analyze_with_claude

# ── Page Config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Discord Sentiment Dashboard",
    page_icon="💬",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os
DATA_DIR = Path(st.secrets.get("OUTPUT_DIR", os.environ.get("OUTPUT_DIR", "./data")))
ROLLING_CSV = DATA_DIR / "messages_all.csv"

# ── Helpers ────────────────────────────────────────────────────────────────────
def load_and_analyze(csv_path: Path, api_key: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path)
    df["timestamp"] = pd.to_datetime(df["timestamp"], utc=True)
    df = df[df["content"].notna() & (df["content"].str.strip() != "")]

    cache_path = csv_path.parent / (csv_path.stem + "_sentiment_cache.csv")
    df = analyze_with_claude(df, api_key, cache_path)

    df["date"] = df["timestamp"].dt.date
    df["week"]  = df["timestamp"].dt.to_period("W").dt.start_time.dt.date

    return df


def sentiment_color(val: float) -> str:
    if val > 0.05:
        return "#1d9e75"
    if val < -0.05:
        return "#d85a30"
    return "#888780"


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Filters")

    try:
        api_key = st.secrets["ANTHROPIC_API_KEY"]
    except KeyError:
        st.error("ANTHROPIC_API_KEY not found. Add it to .streamlit/secrets.toml locally, or to your app's secrets in Streamlit Community Cloud.")
        st.stop()

    if not ROLLING_CSV.exists():
        st.error(f"No data file found at:\n`{ROLLING_CSV}`\n\nRun the bot first.")
        st.stop()

    df_full = load_and_analyze(ROLLING_CSV, api_key)

    min_date = df_full["timestamp"].min().date()
    max_date = df_full["timestamp"].max().date()

    date_range = st.date_input(
        "Date range",
        value=(max_date - timedelta(days=30), max_date),
        min_value=min_date,
        max_value=max_date,
    )

    all_channels = sorted(df_full["channel"].unique())
    selected_channels = st.multiselect(
        "Channels",
        options=all_channels,
        default=all_channels,
    )

    min_length = st.slider("Min message length (chars)", 0, 200, 10)

    st.markdown("---")
    st.caption(f"Last refreshed: {datetime.now().strftime('%H:%M:%S')}")
    if st.button("🔄 Refresh data"):
        st.cache_data.clear()
        st.rerun()


# ── Filter ─────────────────────────────────────────────────────────────────────
start, end = (date_range[0], date_range[1]) if len(date_range) == 2 else (min_date, max_date)

df = df_full[
    (df_full["timestamp"].dt.date >= start)
    & (df_full["timestamp"].dt.date <= end)
    & (df_full["channel"].isin(selected_channels))
    & (df_full["content"].str.len() >= min_length)
].copy()


# ── Header ─────────────────────────────────────────────────────────────────────
st.title("💬 Discord Sentiment Dashboard")
st.caption(f"Showing **{len(df):,}** messages · {start} → {end}")


# ── KPI Row ────────────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)

avg_compound = df["compound"].mean()
pct_pos  = (df["sentiment_label"] == "Positive").mean() * 100
pct_neg  = (df["sentiment_label"] == "Negative").mean() * 100
pct_neu  = (df["sentiment_label"] == "Neutral").mean() * 100
active_authors = df["author_id"].nunique()

k1.metric("Avg sentiment", f"{avg_compound:+.3f}", help="VADER compound score (−1 to +1)")
k2.metric("Positive", f"{pct_pos:.1f}%")
k3.metric("Neutral",  f"{pct_neu:.1f}%")
k4.metric("Negative", f"{pct_neg:.1f}%")
k5.metric("Active authors", f"{active_authors:,}")

st.markdown("---")


# ── Sentiment over time ────────────────────────────────────────────────────────
col_left, col_right = st.columns([2, 1])

with col_left:
    st.subheader("Sentiment over time")

    granularity = st.radio("Granularity", ["Daily", "Weekly"], horizontal=True)
    group_col = "date" if granularity == "Daily" else "week"

    ts = (
        df.groupby(group_col)["compound"]
        .agg(["mean", "count", "std"])
        .reset_index()
        .rename(columns={"mean": "avg_compound", "count": "messages", "std": "std_compound"})
    )
    ts[group_col] = pd.to_datetime(ts[group_col])

    fig_ts = go.Figure()

    # Confidence band
    fig_ts.add_trace(go.Scatter(
        x=pd.concat([ts[group_col], ts[group_col][::-1]]),
        y=pd.concat([
            ts["avg_compound"] + ts["std_compound"].fillna(0),
            (ts["avg_compound"] - ts["std_compound"].fillna(0))[::-1],
        ]),
        fill="toself",
        fillcolor="rgba(83, 74, 183, 0.10)",
        line=dict(color="rgba(0,0,0,0)"),
        showlegend=False,
        hoverinfo="skip",
    ))

    fig_ts.add_trace(go.Scatter(
        x=ts[group_col],
        y=ts["avg_compound"],
        mode="lines+markers",
        line=dict(color="#534AB7", width=2),
        marker=dict(size=5),
        name="Avg compound",
        hovertemplate="%{x|%b %d}<br>Score: %{y:.3f}<extra></extra>",
    ))

    fig_ts.add_hline(y=0, line_dash="dot", line_color="#888780", line_width=1)
    fig_ts.update_layout(
        height=280, margin=dict(l=0, r=0, t=10, b=0),
        yaxis_title="Compound score",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(gridcolor="rgba(128,128,128,0.15)", zeroline=False),
        xaxis=dict(gridcolor="rgba(128,128,128,0.08)"),
    )
    st.plotly_chart(fig_ts, use_container_width=True)

with col_right:
    st.subheader("Sentiment split")

    label_counts = df["sentiment_label"].value_counts()
    fig_pie = go.Figure(go.Pie(
        labels=label_counts.index,
        values=label_counts.values,
        marker_colors=["#1d9e75", "#888780", "#d85a30"],
        hole=0.55,
        textinfo="percent+label",
        hovertemplate="%{label}: %{value:,} messages<extra></extra>",
    ))
    fig_pie.update_layout(
        height=280, margin=dict(l=0, r=0, t=10, b=0),
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_pie, use_container_width=True)


# ── Per-channel breakdown ──────────────────────────────────────────────────────
st.subheader("Channel breakdown")

channel_stats = (
    df.groupby("channel")
    .agg(
        messages=("message_id", "count"),
        avg_sentiment=("compound", "mean"),
        pct_positive=("sentiment_label", lambda x: (x == "Positive").mean() * 100),
        pct_negative=("sentiment_label", lambda x: (x == "Negative").mean() * 100),
        authors=("author_id", "nunique"),
    )
    .reset_index()
    .sort_values("avg_sentiment", ascending=False)
)

fig_bar = go.Figure(go.Bar(
    x=channel_stats["channel"],
    y=channel_stats["avg_sentiment"],
    marker_color=[sentiment_color(v) for v in channel_stats["avg_sentiment"]],
    hovertemplate=(
        "#%{x}<br>"
        "Avg score: %{y:.3f}<br>"
        "<extra></extra>"
    ),
))
fig_bar.add_hline(y=0, line_dash="dot", line_color="#888780", line_width=1)
fig_bar.update_layout(
    height=260, margin=dict(l=0, r=0, t=10, b=0),
    yaxis_title="Avg compound score",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
    xaxis=dict(gridcolor="rgba(0,0,0,0)"),
)
st.plotly_chart(fig_bar, use_container_width=True)

# Table
st.dataframe(
    channel_stats.style.format({
        "avg_sentiment": "{:+.3f}",
        "pct_positive": "{:.1f}%",
        "pct_negative": "{:.1f}%",
    }).background_gradient(
        subset=["avg_sentiment"], cmap="RdYlGn", vmin=-0.5, vmax=0.5
    ),
    use_container_width=True,
    hide_index=True,
)


# ── Top & bottom messages ──────────────────────────────────────────────────────
st.markdown("---")
t_col, b_col = st.columns(2)

with t_col:
    st.subheader("😊 Most positive messages")
    top = df.nlargest(10, "compound")[["timestamp", "channel", "author_name", "content", "compound"]]
    top["compound"] = top["compound"].map("{:+.3f}".format)
    top["timestamp"] = top["timestamp"].dt.strftime("%b %d %H:%M")
    st.dataframe(top, use_container_width=True, hide_index=True)

with b_col:
    st.subheader("😠 Most negative messages")
    bot = df.nsmallest(10, "compound")[["timestamp", "channel", "author_name", "content", "compound"]]
    bot["compound"] = bot["compound"].map("{:+.3f}".format)
    bot["timestamp"] = bot["timestamp"].dt.strftime("%b %d %H:%M")
    st.dataframe(bot, use_container_width=True, hide_index=True)


# ── Author leaderboard ─────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Author sentiment leaderboard")

min_messages = st.slider("Minimum messages to appear", 1, 50, 5)

author_stats = (
    df.groupby(["author_id", "author_name"])
    .agg(
        messages=("message_id", "count"),
        avg_sentiment=("compound", "mean"),
        pct_positive=("sentiment_label", lambda x: (x == "Positive").mean() * 100),
    )
    .reset_index()
    .query(f"messages >= {min_messages}")
    .sort_values("avg_sentiment", ascending=False)
    .drop(columns="author_id")
)

fig_authors = px.scatter(
    author_stats,
    x="messages",
    y="avg_sentiment",
    size="messages",
    color="avg_sentiment",
    color_continuous_scale=["#d85a30", "#888780", "#1d9e75"],
    range_color=[-0.5, 0.5],
    hover_name="author_name",
    hover_data={"messages": True, "avg_sentiment": ":.3f", "pct_positive": ":.1f"},
    labels={"messages": "Message count", "avg_sentiment": "Avg sentiment"},
    height=320,
)
fig_authors.update_layout(
    margin=dict(l=0, r=0, t=10, b=0),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    coloraxis_showscale=False,
    yaxis=dict(gridcolor="rgba(128,128,128,0.15)"),
    xaxis=dict(gridcolor="rgba(128,128,128,0.08)"),
)
st.plotly_chart(fig_authors, use_container_width=True)


# ── Raw data explorer ──────────────────────────────────────────────────────────
with st.expander("🔍 Raw message explorer"):
    search = st.text_input("Search message content")
    show_df = df[df["content"].str.contains(search, case=False, na=False)] if search else df
    st.dataframe(
        show_df[["timestamp", "channel", "author_name", "content", "compound", "sentiment_label"]]
        .sort_values("timestamp", ascending=False)
        .head(500),
        use_container_width=True,
        hide_index=True,
    )
    st.caption(f"Showing up to 500 of {len(show_df):,} matching messages")