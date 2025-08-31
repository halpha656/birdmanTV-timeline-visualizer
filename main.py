# MIT License

# Copyright (c) 2025 halpha656

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

"""
まとめタイムライン生成スクリプト

Usage:
    python generate_timeline_all.py --excel input.xlsx --sheets 2019 2021 2022 2023 --out timeline.png

必要ライブラリ:
    pandas matplotlib openpyxl
"""

import argparse
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib import font_manager as fm
import os


# Robust Japanese font setup: try to register Noto CJK if present, else provide sane fallbacks
def _setup_japanese_font():
    candidates_paths = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf",
        "/usr/share/fonts/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/noto/NotoSansCJKjp-Regular.otf",
    ]
    candidates_names = [
        "Noto Sans CJK JP",
        "Noto Sans CJK",
        "Noto Sans JP",
        "Source Han Sans JP",
        "IPAGothic",
        "TakaoGothic",
        "VL Gothic",
    ]

    # 1) If a known file exists, register it explicitly so matplotlib can find it by name
    for p in candidates_paths:
        if os.path.isfile(p):
            try:
                fm.fontManager.addfont(p)
                name = fm.FontProperties(fname=p).get_name()
                plt.rcParams["font.family"] = [name]
                return
            except Exception:
                pass

    # 2) Try to detect an installed JP-capable font by name
    try:
        installed = {}
        for f in fm.findSystemFonts(fontpaths=None, fontext="ttf") + fm.findSystemFonts(
            fontpaths=None, fontext="otf"
        ):
            try:
                n = fm.FontProperties(fname=f).get_name()
                installed.setdefault(n, f)
            except Exception:
                continue
        for name in candidates_names:
            if name in installed:
                plt.rcParams["font.family"] = [name]
                return
    except Exception:
        pass

    # 3) Fallback: prefer these families if available on the system
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = candidates_names + list(
        plt.rcParams.get("font.sans-serif", [])
    )


_setup_japanese_font()
mpl.rcParams["axes.unicode_minus"] = False


def get_color(dept, note):
    note, dept = note or "", dept or ""
    if "人力" in dept:
        if "VTR" in note:
            return "#2fa0fc"
        else:
            return "#b3e9f8"
    elif "滑空" in dept:
        if "VTR" in note:
            return "#00d65b"
        else:
            return "#cbf266"
    else:
        return "#c8c8cb"


def process_sheet(xls_path, sheet, include_cm: bool):
    df = pd.read_excel(xls_path, sheet_name=sheet, header=0)
    for c in ["尺", "チーム", "部門", "備考"]:
        if c not in df.columns:
            df[c] = ""
    df["尺"] = pd.to_numeric(df["尺"], errors="coerce")
    df = df.dropna(subset=["尺"])
    df = df[df["尺"] > 0]
    # Normalize index early so position-based access is stable
    df = df.reset_index(drop=True)
    df["チーム"] = df["チーム"].fillna("").astype(str)
    df["部門"] = df["部門"].fillna("").astype(str)
    df["備考"] = df["備考"].fillna("").astype(str)
    # CM判定と除外オプション
    mask_cm = (
        df["チーム"].str.contains("CM", na=False)
        | df["備考"].str.contains("CM", na=False)
        | df["部門"].str.contains("CM", na=False)
    )
    df["is_cm"] = mask_cm
    if not include_cm:
        df = df[~mask_cm].reset_index(drop=True)
    # start, end
    df["start"] = df["尺"].cumsum() - df["尺"]
    df["end"] = df["start"] + df["尺"]
    df["is_digest"] = df["備考"].str.contains("ダイジェスト", na=False)
    return df


def generate(excel, sheets, out_path, include_cm: bool):
    data = {s: process_sheet(excel, s, include_cm) for s in sheets}
    row_h = 1.6
    fig, ax = plt.subplots(figsize=(18, row_h * len(sheets) + 2), dpi=150)

    # Helper to place outside labels above and below each row without overlaps
    def draw_row_outside_labels(items, y_base, xmin, xmax):
        # Track vertical extents for this row (include bar edges at least)
        bar_half = 0.35  # matches broken_barh height=0.7
        top_max_y = y_base + bar_half
        bottom_min_y = y_base - bar_half
        if not items:
            return top_max_y, bottom_min_y
        # Sort by anchor to reduce crossings, then split to top/bottom alternately
        items = sorted(items, key=lambda d: d["anchor_x"])
        top_items, bottom_items = [], []
        for i, d in enumerate(items):
            (top_items if i % 2 == 0 else bottom_items).append(d)

        # Geometry: bar half height and gaps
        gap = 0.12  # small gap from bar edge to first label
        lane_spacing = 0.22  # tighter lane spacing

        def place_band(band_items, direction):
            nonlocal top_max_y, bottom_min_y
            if not band_items:
                return
            lane_end = []  # track rightmost x for each lane
            min_gap = 10.0  # seconds between labels
            left_margin, right_margin = 8.0, 12.0
            # base y at which the first lane sits (just outside bar edge)
            base_y = y_base + (bar_half + gap if direction > 0 else -(bar_half + gap))

            for d in band_items:
                w = d["width"]
                # initial preferred x, clamped to axis width
                x0 = max(xmin + left_margin, min(d["pref_x"], xmax - w - right_margin))

                placed = False
                for i, endx in enumerate(lane_end):
                    x_try = max(x0, endx + min_gap)
                    if x_try <= xmax - w - right_margin:
                        y = base_y + direction * (i * lane_spacing)
                        ax.text(
                            x_try,
                            y,
                            d["label"],
                            ha="left",
                            va=("bottom" if direction > 0 else "top"),
                            fontsize=8,
                            style=(d["style"] or "normal"),
                        )
                        ax.annotate(
                            "",
                            xy=(
                                d["anchor_x"],
                                y_base + (bar_half if direction > 0 else -bar_half),
                            ),
                            xytext=(x_try, y),
                            arrowprops=dict(
                                arrowstyle="-",
                                lw=0.6,
                                color="#666",
                                connectionstyle="angle3,angleA=0,angleB=90",
                                shrinkA=0,
                                shrinkB=0,
                            ),
                        )
                        lane_end[i] = x_try + w
                        # update row extents
                        if direction > 0:
                            top_max_y = max(top_max_y, y)
                        else:
                            bottom_min_y = min(bottom_min_y, y)
                        placed = True
                        break
                if not placed:
                    i = len(lane_end)
                    x_try = x0
                    y = base_y + direction * (i * lane_spacing)
                    ax.text(
                        x_try,
                        y,
                        d["label"],
                        ha="left",
                        va=("bottom" if direction > 0 else "top"),
                        fontsize=8,
                        style=(d["style"] or "normal"),
                    )
                    ax.annotate(
                        "",
                        xy=(
                            d["anchor_x"],
                            y_base + (bar_half if direction > 0 else -bar_half),
                        ),
                        xytext=(x_try, y),
                        arrowprops=dict(
                            arrowstyle="-",
                            lw=0.6,
                            color="#666",
                            connectionstyle="angle3,angleA=0,angleB=90",
                            shrinkA=0,
                            shrinkB=0,
                        ),
                    )
                    lane_end.append(x_try + w)
                    # update row extents
                    if direction > 0:
                        top_max_y = max(top_max_y, y)
                    else:
                        bottom_min_y = min(bottom_min_y, y)

        # place above and below
        place_band(top_items, direction=+1)
        place_band(bottom_items, direction=-1)
        return top_max_y, bottom_min_y

    row_centers = []
    y0 = 0
    for sheet in sheets:
        df = data[sheet]
        # その行の外側ラベルを貯めて後で整列して描画
        row_outside = []  # dict(label, anchor_x, pref_x, width, style)
        # 棒を描画
        for _, r in df.iterrows():
            fc = get_color(r["部門"], r["備考"])
            if bool(r.get("is_cm", False)):
                fc = "#9a9a9a"  # 少し濃い灰色
            ax.broken_barh(
                [(r["start"], r["尺"])],
                (y0 - 0.35, 0.7),
                facecolors=fc,
                edgecolors="k",
                linewidth=0.6,
            )
            # CMはバー内に縦書き風の注釈 C\nM を付ける（外側注釈はしない）
            if include_cm and bool(r.get("is_cm", False)):
                xc = (r["start"] + r["end"]) / 2
                ax.text(
                    xc,
                    y0,
                    "C\nM",
                    ha="center",
                    va="center",
                    fontsize=8,
                    color="#000000",
                )
        # シート名を左端に
        ax.text(-50, y0, sheet, ha="right", va="center", fontsize=10, weight="bold")

        # チーム名（同一連続はまとめ）
        merged, cur_team, cur_s, cur_e = [], None, None, None
        for _, r in df.iterrows():
            if r["is_digest"]:
                if cur_team is not None:
                    merged.append((cur_team, cur_s, cur_e))
                    cur_team = None
                continue
            # CM は注釈しない＆シーケンスを分断する
            if bool(r.get("is_cm", False)):
                if cur_team is not None:
                    merged.append((cur_team, cur_s, cur_e))
                    cur_team = None
                continue
            if r["チーム"] == cur_team:
                cur_e = r["end"]
            else:
                if cur_team is not None:
                    merged.append((cur_team, cur_s, cur_e))
                cur_team, cur_s, cur_e = r["チーム"], r["start"], r["end"]
        if cur_team:
            merged.append((cur_team, cur_s, cur_e))

        # ダイジェストまとめ
        digest_groups = []
        i = 0
        n = len(df)
        while i < n:
            if bool(df.iloc[i]["is_digest"]):
                j = i
                while j + 1 < n and bool(df.iloc[j + 1]["is_digest"]):
                    j += 1
                digest_groups.append(
                    (df.iloc[i]["start"], df.iloc[j]["end"], j - i + 1)
                )
                i = j + 1
            else:
                i += 1

        # ラベル配置
        # 固定横軸幅（秒）に合わせて必要なテキスト幅を見積もる
        axis_len = 7200
        fig.canvas.draw()
        fig_w_px = fig.get_size_inches()[0] * fig.dpi
        sec_per_px = axis_len / fig_w_px

        def need_sec(label, fontsize=8):
            t = ax.text(0, 0, label, fontsize=fontsize, alpha=0.0)
            fig.canvas.draw()
            bb = t.get_window_extent(renderer=fig.canvas.get_renderer())
            t.remove()
            return bb.width * sec_per_px + 6

        # 通常チーム
        for team, s, e in merged:
            if not team.strip():
                continue
            L, xc = e - s, (s + e) / 2
            if L > need_sec(team):
                ax.text(xc, y0, team, ha="center", va="center", fontsize=8)
            else:
                row_outside.append(
                    {
                        "label": team,
                        "anchor_x": xc,
                        "pref_x": e + 6,  # 棒の右端から少し右
                        "width": need_sec(team),
                        "style": None,
                    }
                )

        # ダイジェスト
        for s, e, N in digest_groups:
            label = f"ダイジェスト: 計{N}チーム"
            L, xc = e - s, (s + e) / 2
            if L > need_sec(label):
                ax.text(
                    xc, y0, label, ha="center", va="center", fontsize=8, style="italic"
                )
            else:
                row_outside.append(
                    {
                        "label": label,
                        "anchor_x": xc,
                        "pref_x": e + 6,
                        "width": need_sec(label),
                        "style": "italic",
                    }
                )

        # 行ごとの外側ラベルを重ならないように段積み配置して描画
        row_top, row_bottom = draw_row_outside_labels(
            row_outside, y0, xmin=0, xmax=7200
        )
        # track global extents
        try:
            global_top = max(global_top, row_top)
            global_bottom = min(global_bottom, row_bottom)
        except NameError:
            # initialize on first row
            global_top = row_top
            global_bottom = row_bottom
        # stash into axis for later use if needed
        ax._ylim_tracker = (global_bottom, global_top)
        row_centers.append(y0)
        y0 += row_h

    # 外側ラベルは行ごとに描画済み

    # 横軸は 0〜7200 秒で固定、600 秒刻みの目盛りと縦グリッド
    ax.set_xlim(0, 7200)
    try:
        ax.xaxis.set_major_locator(mpl.ticker.MultipleLocator(600))
    except Exception:
        pass
    ax.set_axisbelow(True)
    ax.grid(
        axis="x",
        which="major",
        linestyle="--",
        linewidth=0.6,
        color="#CCCCCC",
        alpha=0.8,
    )
    # Compute tight y-limits from tracked extents with small padding
    ylim_tracker = getattr(ax, "_ylim_tracker", None)
    if ylim_tracker is not None:
        gb, gt = ylim_tracker
        pad_top, pad_bottom = 0.25, 0.25
        ax.set_ylim(gb - pad_bottom, gt + pad_top)
    else:
        ax.set_ylim(-0.8, y0 + 0.8)
    ax.set_yticks([])
    ax.set_xlabel("秒")
    ax.set_title("放送時間と尺のまとめ")
    plt.tight_layout()
    plt.savefig(out_path, dpi=200)
    print("Saved:", out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--sheets", nargs="+", required=True)
    ap.add_argument("--out", default="timeline.png")
    ap.add_argument(
        "--include-cm", action="store_true", help="CMを含める（注釈は付けない）"
    )
    args = ap.parse_args()
    generate(args.excel, args.sheets, args.out, include_cm=args.include_cm)


if __name__ == "__main__":
    main()
