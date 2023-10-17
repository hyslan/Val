import asyncio
from transact_zsbmm216 import novasp


loop = asyncio.get_event_loop()
loop.run_until_complete(novasp("123"))
