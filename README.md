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

`docs/` 中的核心数据由浏览器直接读取：

- `上古音韵数据.json`
- `上古汉语音节表.xlsx`
- `陳靖《兩周古文字編注》索引.xlsx`
- `ids_lv0.txt`

这些文件都已复制到 `docs/`，应随仓库一起部署到 GitHub Pages。
