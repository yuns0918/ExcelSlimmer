from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


def create_icon(base_dir: Path) -> None:
    """Create a simple ExcelSlimmer icon.

    투명 배경 위에 초록색 엑셀 파일 모양(접힌 모서리)과 중앙의 흰색 X,
    그리고 하단 모서리의 작은 반짝임 포인트를 그린다.
    """

    size = 256

    # 투명 배경 캔버스
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 초록색 파일 카드 (살짝 위로 치우친 라운드 사각형)
    card_margin_x = 52
    card_margin_top = 34
    card_margin_bottom = 46
    card_rect = (
        card_margin_x,
        card_margin_top,
        size - card_margin_x,
        size - card_margin_bottom,
    )
    card_color = "#1D6F42"  # 엑셀 그린 톤
    draw.rounded_rectangle(card_rect, radius=32, fill=card_color)

    # 접힌 모서리 (오른쪽 위)
    fold_size = 40
    fx2, fy1 = card_rect[2], card_rect[1]
    fold = [
        (fx2 - fold_size, fy1),
        (fx2, fy1),
        (fx2, fy1 + fold_size),
    ]
    draw.polygon(fold, fill="#2FA865")
    # 접힌 부분의 대각선 라인
    draw.line(
        (fx2 - fold_size, fy1, fx2, fy1 + fold_size),
        fill="#166534",
        width=2,
    )

    # 파일 안쪽 가운데 큰 흰색 X (엑셀 느낌)
    x_color = "#FFFFFF"
    x_margin_x = 34
    x_margin_y = 40
    x_rect = (
        card_rect[0] + x_margin_x,
        card_rect[1] + x_margin_y,
        card_rect[2] - x_margin_x,
        card_rect[3] - x_margin_y,
    )
    draw.line(
        (x_rect[0], x_rect[1], x_rect[2], x_rect[3]),
        fill=x_color,
        width=20,
    )
    draw.line(
        (x_rect[0], x_rect[3], x_rect[2], x_rect[1]),
        fill=x_color,
        width=20,
    )

    # 우하단에 슬리머/정리 느낌의 작은 별 모양 포인트
    star_cx = card_rect[2] - 36
    star_cy = card_rect[3] - 40
    star_color = "#FFFFFF"  # 흰색 반짝임 포인트
    star_points = [
        (star_cx, star_cy - 10),
        (star_cx + 3, star_cy - 3),
        (star_cx + 10, star_cy),
        (star_cx + 3, star_cy + 3),
        (star_cx, star_cy + 10),
        (star_cx - 3, star_cy + 3),
        (star_cx - 10, star_cy),
        (star_cx - 3, star_cy - 3),
    ]
    draw.polygon(star_points, fill=star_color)

    # PNG/ICO로 저장
    png_path = base_dir / "ExcelSlimmer.png"
    ico_path = base_dir / "ExcelSlimmer.ico"

    img.save(png_path, format="PNG")
    # 여러 해상도를 포함한 ICO 생성
    img.save(
        ico_path,
        format="ICO",
        sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
    )

    print(f"Saved PNG: {png_path}")
    print(f"Saved ICO: {ico_path}")


if __name__ == "__main__":
    create_icon(Path(__file__).resolve().parent)
