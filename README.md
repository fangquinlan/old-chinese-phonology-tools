# 上古汉语音韵工具

推荐仓库名：`old-chinese-phonology-tools`

这是一个可直接部署到 GitHub Pages 的静态网页工具，用于上古汉语音韵文本处理。发布目录是 `docs/`，不需要后端服务。

## 功能

- 从粘贴文本中提取汉字读音并导出 CSV。
- 将原文转换为上古拟音文本。
- 从《上古汉语音节表.xlsx》读取释义与注释并导出。
- 按《上古汉语音节表.xlsx》与陳靖《兩周古文字編注》索引提取同諧聲域字表。

## GitHub Pages 部署

1. 在 GitHub 创建仓库：`old-chinese-phonology-tools`。
2. 把本目录推送到仓库，网页发布文件已经整理在 `docs/` 目录。
3. 打开仓库 `Settings` -> `Pages`。
4. `Source` 选择 `Deploy from a branch`，分支选择 `main`，目录选择 `/docs`。
5. 保存后等待 GitHub Pages 构建完成，访问地址通常是：

```text
https://<你的用户名>.github.io/old-chinese-phonology-tools/
```

## 本地预览

直接双击 HTML 时，浏览器可能会阻止读取本地数据文件。建议用本地 HTTP 服务预览：

```powershell
python tools\open-app-server.py
```

或：

```powershell
python -m http.server 8081 -d docs
```

然后访问 `http://localhost:8081/`。运行 `python tools\open-app-server.py` 会预览根目录的同版页面。

## 数据文件

为了让网页保持流畅，浏览器**不再直接解析 `.xlsx` 工作簿**（它们在内存里会膨胀到几十 MB），而是读取由构建脚本预先生成的紧凑 JSON：

- `docs/data/phon.json` —— 提取读音 / 转换拟音用的核心数据，页面启动后即加载（约 2.6 MB，gzip 后约 0.7 MB）。
- `docs/data/gloss.json` —— 释义注释，按需懒加载。
- `docs/data/dict_domain.json` —— 同諧聲域（《上古汉语音节表》），按需懒加载。
- `docs/data/chen_index.json` —— 同諧聲域（陳靖索引），按需懒加载。
- `docs/ids_lv0.txt` —— 导出时读取的字头 IDS 数据。
- `docs/vendor/xlsx.full.min.js` —— 本地内置的 SheetJS，**仅用于导出 XLSX**（不再依赖外部 CDN）。

这些文件由以下源工作簿生成（位于仓库根目录）：`上古汉语音节表.xlsx`、`陳靖《兩周古文字編注》索引.xlsx`。

### 重新生成数据

修改源工作簿后，运行：

```powershell
python tools\build_site_data.py
```

`python tools\open-app-server.py` 启动本地预览时也会在工作簿有更新时自动重建这些文件。

> 备注：`docs/` 下旧的 `上古音韵数据.json` 与两个 `.xlsx` 副本已不再被页面读取，可按需删除以减小部署体积（不影响运行）。
