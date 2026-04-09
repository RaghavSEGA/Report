"""
Discord Message Collector Bot
Runs weekly, collects all messages from all accessible channels,
and exports them to a CSV file for sentiment analysis.
"""

import discord
import asyncio
import csv
import os
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import pandas as pd

# ── Configuration ──────────────────────────────────────────────────────────────
BOT_TOKEN = os.environ.get("DISCORD_BOT_TOKEN", "")
GUILD_ID = int(os.environ.get("DISCORD_GUILD_ID", "0"))  # Your server ID
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "./data"))
LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "7"))  # How far back to fetch

# Channels to skip (e.g. bot-spam, rules, announcements)
SKIP_CHANNELS = set(os.environ.get("SKIP_CHANNELS", "").split(","))

# Schedule: every Monday at 02:00 UTC
SCHEDULE_DAY = os.environ.get("SCHEDULE_DAY", "mon")
SCHEDULE_HOUR = int(os.environ.get("SCHEDULE_HOUR", "2"))
SCHEDULE_MINUTE = int(os.environ.get("SCHEDULE_MINUTE", "0"))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── Bot Setup ──────────────────────────────────────────────────────────────────
intents = discord.Intents.default()
intents.message_content = True
intents.guilds = True

client = discord.Client(intents=intents)
scheduler = AsyncIOScheduler()


# ── Core Collection Logic ──────────────────────────────────────────────────────
async def collect_messages() -> list[dict]:
    """
    Fetch all messages from the past LOOKBACK_DAYS across every
    text channel the bot can read in the target guild.
    """
    guild = client.get_guild(GUILD_ID)
    if not guild:
        log.error("Guild %s not found. Check DISCORD_GUILD_ID.", GUILD_ID)
        return []

    since = datetime.now(timezone.utc) - timedelta(days=LOOKBACK_DAYS)
    rows: list[dict] = []

    text_channels = [
        ch for ch in guild.channels
        if isinstance(ch, discord.TextChannel)
        and ch.name not in SKIP_CHANNELS
        and ch.permissions_for(guild.me).read_messages
        and ch.permissions_for(guild.me).read_message_history
    ]

    log.info("Collecting from %d channels since %s", len(text_channels), since.date())

    for channel in text_channels:
        log.info("  → #%s", channel.name)
        try:
            async for message in channel.history(after=since, limit=None, oldest_first=True):
                if message.author.bot:
                    continue  # Skip bot messages
                rows.append({
                    "message_id":   str(message.id),
                    "channel":      channel.name,
                    "author_id":    str(message.author.id),
                    "author_name":  str(message.author.display_name),
                    "content":      message.content,
                    "timestamp":    message.created_at.isoformat(),
                    "reactions":    sum(r.count for r in message.reactions),
                    "attachments":  len(message.attachments),
                    "reply_to":     str(message.reference.message_id) if message.reference else "",
                })
        except discord.Forbidden:
            log.warning("  No permission to read #%s, skipping.", channel.name)
        except discord.HTTPException as exc:
            log.error("  HTTP error on #%s: %s", channel.name, exc)

        # Polite pause between channels to stay well within rate limits
        await asyncio.sleep(1)

    log.info("Collected %d messages total.", len(rows))
    return rows


def save_csv(rows: list[dict]) -> Path:
    """
    Write rows to a timestamped CSV and also update/append to a
    rolling 'messages_all.csv' that Streamlit reads.
    """
    if not rows:
        log.warning("No messages to save.")
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    snapshot_path = OUTPUT_DIR / f"messages_{timestamp}.csv"

    df = pd.DataFrame(rows)
    df["timestamp"] = pd.to_datetime(df["timestamp"])
    df = df.sort_values("timestamp")

    # Save this week's snapshot
    df.to_csv(snapshot_path, index=False)
    log.info("Saved snapshot → %s (%d rows)", snapshot_path, len(df))

    # Rolling file: append new messages, deduplicate by message_id
    rolling_path = OUTPUT_DIR / "messages_all.csv"
    if rolling_path.exists():
        existing = pd.read_csv(rolling_path, dtype=str)
        combined = pd.concat([existing, df.astype(str)], ignore_index=True)
        combined = combined.drop_duplicates(subset=["message_id"])
        combined.to_csv(rolling_path, index=False)
        log.info("Updated rolling file → %s (%d total rows)", rolling_path, len(combined))
    else:
        df.to_csv(rolling_path, index=False)
        log.info("Created rolling file → %s", rolling_path)

    return snapshot_path


# ── Scheduled Job ──────────────────────────────────────────────────────────────
async def weekly_job():
    log.info("=== Weekly collection job started ===")
    rows = await collect_messages()
    save_csv(rows)
    log.info("=== Weekly collection job complete ===")


# ── Bot Events ─────────────────────────────────────────────────────────────────
@client.event
async def on_ready():
    log.info("Logged in as %s (ID: %s)", client.user, client.user.id)
    log.info("Watching guild ID: %s", GUILD_ID)

    scheduler.add_job(
        weekly_job,
        CronTrigger(day_of_week=SCHEDULE_DAY, hour=SCHEDULE_HOUR, minute=SCHEDULE_MINUTE),
        id="weekly_collection",
        replace_existing=True,
    )
    scheduler.start()

    log.info(
        "Scheduler started. Next run: %s",
        scheduler.get_job("weekly_collection").next_run_time,
    )

    # Uncomment to run immediately on startup (useful for first-time setup):
    # await weekly_job()


@client.event
async def on_message(message: discord.Message):
    if message.content == "!collect" and message.author.guild_permissions.administrator:
        await message.reply("Starting manual collection…")
        await weekly_job()
        await message.reply("Done! CSV updated.")


# ── Entry Point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not BOT_TOKEN:
        raise SystemExit("DISCORD_BOT_TOKEN environment variable is not set.")
    if not GUILD_ID:
        raise SystemExit("DISCORD_GUILD_ID environment variable is not set.")

    client.run(BOT_TOKEN)
