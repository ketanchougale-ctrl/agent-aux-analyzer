"""Generate simple bar-chart PNG icons for the Chrome extension (no dependencies)."""
import struct, zlib, os

def make_chunk(ctype: bytes, data: bytes) -> bytes:
    c = ctype + data
    return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

def make_png(size: int) -> bytes:
    BG  = (21,  101, 192)   # #1565C0 blue
    BAR = (255, 255, 255)   # white bars

    pad      = max(2, size // 10)
    bar_w    = max(2, size // 7)
    gap      = max(1, size // 12)
    heights  = [int(size * 0.68), int(size * 0.44), int(size * 0.56)]

    starts = []
    x = pad
    for _ in heights:
        starts.append(x)
        x += bar_w + gap

    raw = b""
    for y in range(size):
        raw += b"\x00"               # filter: none
        for x in range(size):
            color = BG
            for bx, bh in zip(starts, heights):
                if bx <= x < bx + bar_w and y >= size - pad - bh:
                    color = BAR
                    break
            raw += bytes(color)

    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    png  = b"\x89PNG\r\n\x1a\n"
    png += make_chunk(b"IHDR", ihdr)
    png += make_chunk(b"IDAT", zlib.compress(raw, 9))
    png += make_chunk(b"IEND", b"")
    return png

os.makedirs("icons", exist_ok=True)
for s in [16, 48, 128]:
    path = f"icons/icon{s}.png"
    with open(path, "wb") as f:
        f.write(make_png(s))
    print(f"[OK]  Created {path}  ({s}x{s} px)")

print("\nAll icons created. Load the chrome-extension folder in Chrome.")
