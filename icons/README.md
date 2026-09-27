# icons/

PWA／主畫面用的圖示。全部由 `index.html` 同一套視覺（藍底＋白文件＋W）產生：

```python
# 需要 Pillow
from PIL import Image, ImageDraw
def icon(size, bg=(37,99,235), fg=(255,255,255)):   # 藍底圓角＋白色文件＋藍色 W
    ...
```

| 檔案 | 尺寸 | 用途 |
|---|---|---|
| `icon-192.png` / `icon-512.png` | 192／512 | manifest `purpose: any` |
| `icon-192-maskable.png` / `icon-512-maskable.png` | 192／512 | manifest `purpose: maskable`（內容縮至 62% 留安全區） |
| `apple-touch-icon.png` | 180 | iOS 加到主畫面 |
| `favicon-32.png` / `favicon-16.png` | 32／16 | 瀏覽器分頁圖示 |

換圖示時記得同步 `manifest.webmanifest` 的 `icons` 與 `index.html` `<head>` 的連結，
並把 `changelog.json` 的 `version` 換掉（Service Worker 的快取名稱就是它，換了才會重抓）。
