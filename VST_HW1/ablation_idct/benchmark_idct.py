import time
from pathlib import Path

import numpy as np

from codec_parameters import DCT_BASIS, HEADER, Q, ZIGZAG
from decoder import BitReader, read_block, direct_idct

def load_blocks(path):
    raw = Path(path).read_bytes()
    valid_bits = HEADER.unpack(raw[:HEADER.size])[0]
    reader = BitReader(raw[HEADER.size:], valid_bits)
    blocks = []

    for _ in range(4096):
        zigzagged, _ = read_block(reader)
        flat = np.zeros(64, dtype=np.int32)
        flat[ZIGZAG] = zigzagged
        blocks.append(flat.reshape(8, 8) * Q)

    assert reader.pos == valid_bits

    return blocks

def two_stage_idct(coeffs):
    after_columns = DCT_BASIS.T @ coeffs
    
    return after_columns @ DCT_BASIS

def measure(function, blocks):
    start = time.perf_counter()

    for coeffs in blocks:
        function(coeffs)

    return time.perf_counter() - start

root = Path(__file__).resolve().parent.parent

for name in ["lena", "mandril"]:
    blocks = load_blocks(root / f"{name}.bin")

    max_error = 0.0

    for coeffs in blocks:
        error = np.max(np.abs(
            two_stage_idct(coeffs) - direct_idct(coeffs)
        ))

        max_error = max(max_error, float(error))

    stage_times = []
    direct_times = []

    for repeat in range(3):
        if repeat % 2 == 0:
            stage_times.append(measure(two_stage_idct, blocks))
            direct_times.append(measure(direct_idct, blocks))
        else:
            direct_times.append(measure(direct_idct, blocks))
            stage_times.append(measure(two_stage_idct, blocks))

    stage = float(np.median(stage_times))
    direct = float(np.median(direct_times))
    bpp  = (root / f"{name}.bin").stat().st_size * 8 / (512 * 512)

    print(f"\nImage: {name}")
    print(f"Maximum IDCT difference: {max_error:.12g}")
    print(f"Two-stage times: {stage_times}")
    print(f"Direct times: {direct_times}")
    print(f"Two-stage median: {stage:.6f} seconds")
    print(f"Direct median: {direct:.6f} seconds")
    print(f"Direct / two-stage time: {direct / stage:.2f}")
    print(f"BPP for both methods: {bpp:.6f}")