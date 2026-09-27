# vendor/

這裡放的是**按下功能按鈕才載入**的第三方函式庫。改用 repo 內的本機檔案、不再連 CDN，原因：

- 這個工具的賣點是「檔案在本機瀏覽器處理」，但原本「載入 Word」與「轉換 PDF」都必須先連 jsDelivr。
  沒網路、CDN 被公司網路擋掉、或該版本被下架時，這兩個主要功能會直接失效。
- 本機檔案沒有第三方請求，也少一個隱私與可用性的變數。
- 首次開啟頁面仍然只下載 `index.html`（約 28KB），這兩個檔案是延後載入的。

| 檔案 | 版本 | 用途 | 授權 |
|---|---|---|---|
| `mammoth.browser.min.js` | mammoth 1.12.0 | 解析 `.docx` → HTML | BSD-2-Clause |
| `html2pdf.bundle.min.js` | html2pdf.js 0.10.2（打包 html2canvas 1.4.1、jsPDF 2.5.1） | HTML → PDF（點陣圖） | MIT |

## 怎麼更新

```bash
curl -sL -o vendor/mammoth.browser.min.js    https://cdn.jsdelivr.net/npm/mammoth@1.12.0/mammoth.browser.min.js
curl -sL -o vendor/html2pdf.bundle.min.js    https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.2/dist/html2pdf.bundle.min.js
```

換版本時要一起改 `index.html` 裡的 `LIB` 常數路徑，並在瀏覽器實測「載入 Word」與「轉換 PDF」兩個流程。

**換了這裡的檔案就要把 `sw.js` 的 `VERSION` 加一**：Service Worker 對 `vendor/` 走 cache-first，
版本沒變的話舊快取會繼續被使用，使用者會拿到舊的函式庫。

## 授權全文

### mammoth（BSD-2-Clause）

```
Copyright (c) 2013, Michael Williamson
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```

### html2pdf.js（MIT）

```
The MIT License

Copyright (c) 2017 Erik Koopmans

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

打包在 `html2pdf.bundle.min.js` 內的 html2canvas（MIT，Copyright (c) 2012 Niklas von Hertzen）與
jsPDF（MIT，Copyright (c) 2010-2023 James Hall, yWorks GmbH）同樣為 MIT 授權。
