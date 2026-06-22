import asyncio
import json
import os

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
SESSION_FILE = os.path.join(APP_DIR, "_test_session")


async def main():
    with open(CONFIG_FILE, encoding="utf-8") as f:
        cfg = json.load(f)
    api_id = int(cfg["api_id"])
    api_hash = cfg["api_hash"]
    phone = cfg["phone"]
    print("API ID:", api_id)
    print("Phone:", phone)
    print("Hash:", api_hash[:8] + "...")

    try:
        from telethon import TelegramClient
        from telethon.errors import ApiIdInvalidError, PhoneNumberInvalidError
    except ImportError as e:
        print("FAIL: telethon not installed:", e)
        return

    client = TelegramClient(SESSION_FILE, api_id, api_hash)
    try:
        print("Connecting to Telegram...")
        await client.connect()
        print("OK: connected to Telegram servers")

        authorized = await client.is_user_authorized()
        print("Authorized:", authorized)

        if not authorized:
            print("Sending code request (test)...")
            try:
                sent = await client.send_code_request(phone)
                print("OK: code request sent!")
                print("Code type:", type(sent.type).__name__)
                print("Phone code hash received:", bool(sent.phone_code_hash))
            except ApiIdInvalidError:
                print("FAIL: API ID / API Hash invalid!")
            except PhoneNumberInvalidError:
                print("FAIL: Phone number invalid!")
            except Exception as e:
                print("FAIL on send_code_request:", type(e).__name__, e)
        else:
            me = await client.get_me()
            print("OK: already logged in as", me.first_name, me.username)
    except Exception as e:
        print("FAIL on connect:", type(e).__name__, e)
    finally:
        await client.disconnect()
        print("Disconnected.")


if __name__ == "__main__":
    asyncio.run(main())
