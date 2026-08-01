# Outlook PST 邮件搜索工具

本地小工具：选择一个 `.pst` 文件，按 **主题 / 发件人 / 收件人 / 日期** 组合搜索邮件，
自动生成 Excel 报表（含匹配率），并把匹配到的邮件另存为 `.eml` 文件放入指定文件夹。

**特点：直接解析 PST 文件本身，不需要打开/依赖 Outlook 程序，速度更快、不占用 Outlook 资源。**

---

## 一、给同事使用（最简单方式）

1. 找一台 Windows 电脑，按下面"二、打包"的步骤生成 `PST邮件搜索工具.exe`
2. 把这一个 exe 文件拷贝给所有同事，双击即可运行，**不需要安装 Python 或任何依赖**
3. 使用步骤：
   - 点击"浏览"选择要搜索的 `.pst` 文件
   - 选择结果输出文件夹（默认会在 PST 所在目录下新建"搜索结果"文件夹）
   - 填写搜索条件（主题关键字 / 发件人 / 收件人 / 日期范围），可以只填一个，也可以组合
   - 点击"开始搜索并生成报表"
   - 完成后会弹窗提示：共扫描多少封、匹配多少封、匹配率是多少
   - 输出文件夹里会有：
     - 一个 Excel 报表（`邮件搜索报表_日期时间.xlsx`），第一个表是明细，第二个表是统计汇总
     - 一个"匹配邮件"子文件夹，里面是每封匹配邮件的 `.eml` 文件，双击可直接用 Outlook 打开查看

### ⚠️ 重要提示
如果 `.pst` 文件当前正在 Outlook 中打开着，Windows 会锁定该文件，可能导致读取失败或报错。
请提醒同事：**先关闭 Outlook**，或者先复制一份 PST 文件，再对副本进行搜索。

---

## 二、只有 Mac 电脑？用 GitHub Actions 云端打包（推荐，全程不需要 Windows 电脑）

PyInstaller 不支持"跨平台打包"——在 Mac 上打包只能打出 Mac 版本，不能直接生成 Windows 的 exe。
但可以借助 **GitHub Actions**（GitHub 提供的免费云端环境，包含真实的 Windows 系统）远程帮你打包，
你全程只需要在 Mac 上操作浏览器，最后下载打包好的 exe 文件即可。本工具包里已经附带了配置好的
`.github/workflows/build-windows-exe.yml`，不需要你自己写。

**具体步骤：**

1. 打开 https://github.com ，注册/登录一个账号（免费）
2. 右上角 "+" → "New repository"，随便起个名字（比如 `pst-tool`），Public 或 Private 都可以，创建
3. 进入刚创建的空仓库页面，点击 "uploading an existing file"（或 "Add file" → "Upload files"）
4. 把本工具包里的所有文件和文件夹（包括 `.github` 文件夹本身，连同里面的 `workflows` 子文件夹和
   `build-windows-exe.yml` 一起；也包括 `assets` 文件夹和里面的 `logo.png`，这是界面上显示的公司 logo）
   拖拽上传，然后点击底部绿色的 "Commit changes"
   - 如果网页拖拽上传不认 `.github` 这种隐藏文件夹，改用"三、（进阶）用命令行上传"里的 git 方式即可
5. 上传完成后，点顶部的 **"Actions"** 标签页，会看到一个自动开始运行的任务（因为上传触发了 push）；
   如果没自动开始，点左侧的工作流名字，再点 "Run workflow" 手动触发
6. 等 2-5 分钟，任务跑完变绿色勾 ✅ 后，点进这次运行，页面最下方 "Artifacts" 里有一个
   `PST邮件搜索工具-windows-exe` 的压缩包，点击下载
7. 解压后就是 `PST邮件搜索工具.exe`，这是真正的 Windows 可执行文件，
   直接拷给同事在 Windows 电脑上使用即可，同事不需要装任何东西

### 三、（进阶）如果你熟悉命令行，用 git 上传更省事

```bash
cd pst_tool
git init
git add .
git commit -m "init"
git branch -M main
git remote add origin https://github.com/你的用户名/pst-tool.git
git push -u origin main
```
push 完成后同样去仓库的 "Actions" 页面查看和下载结果。

---

## 四、如果实在想自己打包（有 Windows 电脑或虚拟机的情况）

1. 准备一台 **Windows** 电脑，安装 Python 3.10 或以上版本（勾选"Add to PATH"）
2. 把本文件夹（`pst_search_tool.py`、`requirements.txt`、`build_exe.bat`）拷贝到该电脑
3. 双击运行 `build_exe.bat`（或在命令行里进入该文件夹后执行 `build_exe.bat`）
4. 等待安装依赖、打包完成，生成的文件在 `dist\PST邮件搜索工具.exe`
5. 把这一个 exe 文件发给同事即可，同事不需要装 Python

如果打包时 `pip install libpff-python-windows` 报错（例如没有匹配的 Python 版本预编译包），
可以换成从源码编译版本：`pip install libpff-python`，但这个方式需要额外安装 C 编译工具，
不建议给非技术同事操作，建议由技术人员在一台电脑上编译好之后统一分发 exe。

---

## 五、直接用 Python 运行（不打包成 exe，适合技术人员自己用）

```bash
pip install -r requirements.txt
python pst_search_tool.py
```

---

## 六、已知限制

- 仅支持标准 `.pst` 格式（Outlook 个人文件夹文件）。如果是 `.ost`（Exchange 缓存文件），
  libpff 理论上也能解析，但建议优先用 pst 测试。
- 不支持采用"高级加密"（NDPack 加密）保护的 PST（早期版本 Outlook 生成的普通压缩加密格式没问题）。
- 大文件（几个 G 以上）扫描可能需要几分钟，请耐心等待，界面下方会显示扫描进度。
- 日期筛选依据邮件的"送达时间"（收不到则用"发送时间"），草稿箱等没有这些时间戳的邮件会被日期条件排除。
- 导出的 `.eml` 是标准邮件格式（非 Outlook 私有的 `.msg` 格式），Outlook / 大部分邮件客户端都可以直接双击打开，
  但正文格式、附件目前版本未包含在导出文件中（如需要连附件一起导出，可以告诉我，我再补充这部分功能）。

---

## 七、文件清单

| 文件 | 说明 |
|---|---|
| `pst_search_tool.py` | 主程序（GUI + 搜索 + 导出逻辑） |
| `assets/logo.png` | 界面左上角显示的公司 logo，打包时会一起打进 exe |
| `requirements.txt` | 依赖库清单 |
| `build_exe.bat` | Windows 电脑上一键打包成 exe 的脚本 |
| `.github/workflows/build-windows-exe.yml` | GitHub Actions 云端自动打包配置（Mac 用户推荐用这个） |
| `README.md` | 本说明文档 |
