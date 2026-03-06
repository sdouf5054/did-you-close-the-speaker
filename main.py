"""
Did You Close the Speaker?
Safe shutdown/sleep pipeline that turns off Tapo P115 smart plug before system power actions.
"""

import sys
import asyncio
import argparse
import logging
from pathlib import Path

from config import load_config
from tapo_control import TapoController
from power import shutdown_windows, sleep_windows, restart_windows

LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "dycts.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("dycts")


async def speaker_off(controller: TapoController) -> bool:
    """Turn off the speaker plug. Returns True on success."""
    try:
        await controller.turn_off()
        logger.info("Speaker plug turned OFF successfully.")
        return True
    except Exception as e:
        logger.error(f"Failed to turn off speaker plug: {e}")
        return False


async def speaker_on(controller: TapoController) -> bool:
    """Turn on the speaker plug. Returns True on success."""
    try:
        await controller.turn_on()
        logger.info("Speaker plug turned ON successfully.")
        return True
    except Exception as e:
        logger.error(f"Failed to turn on speaker plug: {e}")
        return False


async def safe_power_action(controller: TapoController, action: str, delay: float, force: bool = False) -> None:
    """
    Turn off speaker, wait, then perform power action.
    
    Args:
        controller: TapoController instance
        action: One of 'shutdown', 'sleep', 'restart'
        delay: Seconds to wait after turning off plug
        force: If True, proceed with power action even if plug control fails
    """
    success = await speaker_off(controller)

    if not success and not force:
        logger.warning("Speaker plug OFF failed. Aborting power action.")
        print("\n⚠️  Failed to turn off speaker plug!")
        print("   Check your network connection and Tapo device status.")
        print(f"   To force {action} anyway, use: dycts {action} --force")
        sys.exit(1)

    if not success and force:
        logger.warning(f"Speaker plug OFF failed, but --force flag set. Proceeding with {action}.")
        print("\n⚠️  Speaker plug OFF failed, but forcing power action...")

    if success and delay > 0:
        logger.info(f"Waiting {delay}s for speaker to safely power down...")
        await asyncio.sleep(delay)

    action_map = {
        "shutdown": shutdown_windows,
        "sleep": sleep_windows,
        "restart": restart_windows,
    }

    logger.info(f"Executing: {action}")
    action_map[action]()


def main():
    parser = argparse.ArgumentParser(
        prog="dycts",
        description="Did You Close the Speaker? — Safe power management for active monitors.",
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # shutdown
    p_shutdown = subparsers.add_parser("shutdown", help="Turn off speaker, then shutdown PC")
    p_shutdown.add_argument("--force", action="store_true", help="Proceed even if plug OFF fails")

    # sleep
    p_sleep = subparsers.add_parser("sleep", help="Turn off speaker, then sleep PC")
    p_sleep.add_argument("--force", action="store_true", help="Proceed even if plug OFF fails")

    # restart
    p_restart = subparsers.add_parser("restart", help="Turn off speaker, then restart PC")
    p_restart.add_argument("--force", action="store_true", help="Proceed even if plug OFF fails")

    # speaker control
    subparsers.add_parser("off", help="Turn off speaker plug only")
    subparsers.add_parser("on", help="Turn on speaker plug only")
    subparsers.add_parser("status", help="Check speaker plug status")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    cfg = load_config()
    controller = TapoController(
        ip=cfg["plug_ip"],
        email=cfg["tapo_email"],
        password=cfg["tapo_password"],
        timeout=cfg.get("timeout_sec", 5),
    )

    if args.command in ("shutdown", "sleep", "restart"):
        asyncio.run(
            safe_power_action(
                controller=controller,
                action=args.command,
                delay=cfg.get("delay_after_power_off_sec", 2),
                force=args.force,
            )
        )
    elif args.command == "off":
        success = asyncio.run(speaker_off(controller))
        if success:
            print("✅ Speaker plug is now OFF.")
        else:
            print("❌ Failed to turn off speaker plug.")
            sys.exit(1)
    elif args.command == "on":
        success = asyncio.run(speaker_on(controller))
        if success:
            print("✅ Speaker plug is now ON.")
        else:
            print("❌ Failed to turn on speaker plug.")
            sys.exit(1)
    elif args.command == "status":
        info = asyncio.run(controller.get_status())
        if info:
            print(f"🔌 Speaker plug status: {info}")
        else:
            print("❌ Failed to get speaker plug status.")
            sys.exit(1)


if __name__ == "__main__":
    main()
