"""Run the existing Telegram bot and its website API in one FadeHost app."""
import asyncio
import os
import threading

import uvicorn
import bot
from web_api import app as api_app


def _run_telegram_bot():
    # python-telegram-bot uses its own asyncio loop; give the worker thread one explicitly.
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        bot.main(run_in_thread=True)
    finally:
        try:
            loop.close()
        except Exception:
            pass


if __name__ == "__main__":
    thread = threading.Thread(target=_run_telegram_bot, name="telegram-bot", daemon=True)
    thread.start()
    port = int(os.getenv("PORT", "8080"))
    print(f"Website API starting on 0.0.0.0:{port}")
    uvicorn.run(api_app, host="0.0.0.0", port=port, log_level="info")
