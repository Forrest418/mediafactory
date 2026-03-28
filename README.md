# MediaFactory

[English](#english) | [中文](#中文)

## English

MediaFactory is a WPF desktop app for managing AI-assisted image planning and image generation workflows.

### Screenshots

| Workspace | Settings |
| --- | --- |
| ![Workspace](./images/2026-03-28_22-22-23.png) | ![Settings](./images/2026-03-28_22-23-31.png) |

| Planning | Output Review |
| --- | --- |
| ![Planning](./images/2026-03-28_22-39-42.png) | ![Output Review](./images/2026-03-28_23-06-22.png) |

<details>
<summary>More screenshots</summary>

![Screenshot 5](./images/2026-03-28_23-07-34.png)
![Screenshot 6](./images/2026-03-28_23-08-07.png)
![Screenshot 7](./images/2026-03-28_23-08-22.png)
![Screenshot 8](./images/2026-03-28_23-08-47.png)
![Screenshot 9](./images/2026-03-28_23-09-49.png)
![Screenshot 10](./images/2026-03-28_23-10-15.png)
![Screenshot 11](./images/2026-03-28_23-11-27.png)

</details>

### What it does

- Manage multiple image production projects in one workspace
- Import reference images for each project
- Generate design plans and image plans with configurable AI providers
- Batch generate images and track task progress
- Review generated outputs in a media library
- Persist projects, presets, and studio state with SQLite

### Workflow

```mermaid
flowchart LR
    A["Import reference images"] --> B["Select scenario and model providers"]
    B --> C["Generate design plan"]
    C --> D["Review and edit markdown"]
    D --> E["Batch generate images"]
    E --> F["Review outputs and export"]
```

### Tech stack

- VB.NET
- WPF
- .NET 8
- SQLite via `Microsoft.Data.Sqlite`

### Project structure

- `MediaFactory.App/`: application source
- `MediaFactory.slnx`: solution file
- `ModelProviders.json`: safe template config committed to the repo

### Configuration

The app looks for config files in this order:

1. `ModelProviders.local.json`
2. `ModelProviders.json`
3. Legacy fallback: `GoogleVertex.local.json`
4. Legacy fallback: `GoogleVertex.json`

Recommended workflow:

1. Keep the committed `ModelProviders.json` as the public template
2. Create your own `ModelProviders.local.json`
3. Put your real API keys and service account values only in the local file

`ModelProviders.local.json` is already ignored by Git.

### Run locally

```powershell
dotnet build .\MediaFactory.slnx
dotnet run --project .\MediaFactory.App\MediaFactory.App.vbproj
```

### Open source notes

- No real API keys should be committed to the repository
- The DevExpress dependency was removed so the project can build with public tooling
- The repository is ready for GitHub publishing

### Roadmap

- Improve provider abstraction so OpenAI-compatible providers can be added with configuration only
- Add repository screenshots and demo media
- Add connection testing for model providers in the settings page
- Improve prompt template customization and preset sharing
- Add packaging and release automation

### Contributing

Contributions are welcome. Please read [CONTRIBUTING.md](./CONTRIBUTING.md) before opening a pull request.

### Security

If you find a security issue or accidentally discover secrets in a local config file, do not open a public issue. Read [SECURITY.md](./SECURITY.md) for the reporting process.

### Publish to GitHub

If `git` is available in your PATH:

```powershell
git init -b main
git add .
git commit -m "Initial open source release"
git remote add origin https://github.com/<your-account>/MediaFactory.git
git push -u origin main
```

If `git` is not in your PATH on this machine, you can use the Visual Studio bundled Git:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" init -b main
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" add .
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" commit -m "Initial open source release"
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" remote add origin https://github.com/<your-account>/MediaFactory.git
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" push -u origin main
```

## 中文

MediaFactory 是一个基于 WPF 的桌面应用，用于管理 AI 辅助的图片规划与图片生成工作流。

### 界面截图

| 工作台 | 系统设置 |
| --- | --- |
| ![工作台](./images/2026-03-28_22-22-23.png) | ![系统设置](./images/2026-03-28_22-23-31.png) |

| 规划编辑 | 结果查看 |
| --- | --- |
| ![规划编辑](./images/2026-03-28_22-39-42.png) | ![结果查看](./images/2026-03-28_23-06-22.png) |

<details>
<summary>更多截图</summary>

![截图 5](./images/2026-03-28_23-07-34.png)
![截图 6](./images/2026-03-28_23-08-07.png)
![截图 7](./images/2026-03-28_23-08-22.png)
![截图 8](./images/2026-03-28_23-08-47.png)
![截图 9](./images/2026-03-28_23-09-49.png)
![截图 10](./images/2026-03-28_23-10-15.png)
![截图 11](./images/2026-03-28_23-11-27.png)

</details>

### 功能说明

- 在一个工作台中管理多个图片生产项目
- 为每个项目导入参考图片
- 使用可配置的 AI 供应商生成设计规划和图片规划
- 批量生成图片并跟踪任务进度
- 在媒体库中查看和管理生成结果
- 使用 SQLite 持久化项目、模板和工作台状态

### 工作流程

```mermaid
flowchart LR
    A["导入参考图片"] --> B["选择场景与模型供应商"]
    B --> C["生成设计规划"]
    C --> D["审核并编辑 Markdown"]
    D --> E["批量生成图片"]
    E --> F["查看结果并导出"]
```

### 技术栈

- VB.NET
- WPF
- .NET 8
- 通过 `Microsoft.Data.Sqlite` 使用 SQLite

### 项目结构

- `MediaFactory.App/`：应用源码
- `MediaFactory.slnx`：解决方案文件
- `ModelProviders.json`：提交到仓库中的安全配置模板

### 配置说明

程序会按以下顺序查找配置文件：

1. `ModelProviders.local.json`
2. `ModelProviders.json`
3. 兼容旧文件：`GoogleVertex.local.json`
4. 兼容旧文件：`GoogleVertex.json`

推荐使用方式：

1. 将仓库中的 `ModelProviders.json` 保留为公开模板
2. 在本地创建自己的 `ModelProviders.local.json`
3. 将真实 API Key 和服务账号信息只放在本地文件中

`ModelProviders.local.json` 已被 Git 忽略，不会提交到仓库。

### 本地运行

```powershell
dotnet build .\MediaFactory.slnx
dotnet run --project .\MediaFactory.App\MediaFactory.App.vbproj
```

### 开源发布说明

- 仓库中不应提交真实 API Key
- 已移除 DevExpress 依赖，可以用公开工具链直接构建
- 当前仓库已经整理到可发布到 GitHub 的状态

### Roadmap

- 继续抽象模型供应商，让 OpenAI 兼容供应商尽量只靠配置接入
- 补充仓库截图和演示素材
- 为系统设置页增加模型连接测试
- 进一步增强提示词模板自定义和模板分享能力
- 增加打包与发布自动化

### 参与贡献

欢迎提交 Issue 和 Pull Request。提交前请先阅读 [CONTRIBUTING.md](./CONTRIBUTING.md)。

### 安全说明

如果你发现安全问题，或者意外看到本地配置文件中的敏感信息，请不要直接提交公开 Issue。请先阅读 [SECURITY.md](./SECURITY.md) 中的处理方式。

### 发布到 GitHub

如果当前机器的 PATH 中已经有 `git`：

```powershell
git init -b main
git add .
git commit -m "Initial open source release"
git remote add origin https://github.com/<your-account>/MediaFactory.git
git push -u origin main
```

如果当前机器的 PATH 中没有 `git`，可以使用 Visual Studio 自带的 Git：

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" init -b main
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" add .
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" commit -m "Initial open source release"
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" remote add origin https://github.com/<your-account>/MediaFactory.git
& "C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe" push -u origin main
```

## License

MIT
