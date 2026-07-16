# 蒲公英爬虫 macOS版 — 构建说明

## 你自己怎么拿到最终文件

### 方法一：GitHub Actions 自动编译（推荐，免费）

1. 在 GitHub 新建一个仓库（公开/私有都行）
2. 把 `蒲公英爬虫_macOS_build/` 下所有文件推到仓库根目录：
   ```
   git init
   git add .
   git commit -m "init"
   git remote add origin https://github.com/你的用户名/随便起个名.git
   git push -u origin main
   ```
3. 推送后，GitHub 自动开始编译（约 3 分钟）
4. 去仓库的 **Actions** 标签页 → 等黄色圆点变绿 ✓
5. 点击完成的构建 → 页面底部 **Artifacts** → 下载 `蒲公英爬虫_macOS`
6. 解压后得到 3 个文件：
   - `蒲公英采集`（主程序，右键→打开）
   - `启动Chrome.command`（Chrome 启动器）
   - `使用说明_mac.txt`

### 方法二：我手动编译后传给你

如果你不想搞 GitHub，把 `蒲公英爬虫_macOS_build/` 文件夹打包发给我，我在一台 Mac 上帮你编译。但 GitHub Actions 更快更省事。

## 产物说明

| 文件 | 说明 |
|------|------|
| `蒲公英爬虫_macOS` | Intel 编译版，Intel + M1/M2/M3 通用（M 系列通过 Rosetta 2 转译） |
| `蒲公英爬虫_macOS_ARM64` | M 系列原生版，速度稍快但非必须，Intel Mac 无法运行 |

客户用 Intel 版即可，5 台设备全兼容。

## 客户使用

1. 解压，双击 `启动Chrome.command`（首次右键→打开）
2. 双击 `蒲公英采集`（首次右键→打开）
3. 选 Excel 文件 → 如果没登录，在 Chrome 窗口登录 蒲公英
4. 等 2 分钟，`达人数据.csv` 生成在同一文件夹
