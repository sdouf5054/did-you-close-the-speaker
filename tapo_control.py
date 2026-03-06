"""
Tapo P115 smart plug controller.
Uses the `tapo` library for local API communication.
"""

import asyncio
from tapo import ApiClient


class TapoController:
    def __init__(self, ip: str, email: str, password: str, timeout: float = 5.0):
        self.ip = ip
        self.email = email
        self.password = password
        self.timeout = timeout

    async def _connect(self):
        """Connect to the Tapo device and return the device handle."""
        client = ApiClient(self.email, self.password)
        device = await client.p115(self.ip)
        return device

    async def turn_off(self) -> None:
        """Turn off the plug with timeout."""
        async with asyncio.timeout(self.timeout):
            device = await self._connect()
            await device.off()

    async def turn_on(self) -> None:
        """Turn on the plug with timeout."""
        async with asyncio.timeout(self.timeout):
            device = await self._connect()
            await device.on()

    async def get_status(self) -> dict | None:
        """Get device info. Returns dict with status or None on failure."""
        try:
            async with asyncio.timeout(self.timeout):
                device = await self._connect()
                info = await device.get_device_info()
                return {
                    "device_on": info.device_on,
                    "friendly_name": getattr(info, "nickname", "P115"),
                    "signal_level": getattr(info, "signal_level", None),
                }
        except Exception:
            return None
