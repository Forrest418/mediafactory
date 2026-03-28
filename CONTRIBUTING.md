# Contributing

[English](#english) | [中文](#中文)

## English

Thanks for contributing to MediaFactory.

### Before you start

- Open an issue first for large changes
- Keep changes focused and easy to review
- Do not commit API keys, tokens, or service-account files
- Prefer configuration and small abstractions over large one-off patches

### Development setup

```powershell
dotnet build .\MediaFactory.slnx
dotnet run --project .\MediaFactory.App\MediaFactory.App.vbproj
```

### Pull request checklist

- Build passes locally
- No secrets were added to tracked files
- README or config template is updated when behavior changes
- UI changes were checked in both English and Chinese when applicable

## 中文

感谢你为 MediaFactory 做贡献。

### 开始之前

- 如果改动较大，建议先开 Issue 讨论
- 尽量让改动聚焦，便于评审
- 不要提交 API Key、Token 或服务账号文件
- 优先使用配置和小范围抽象，避免一次性的大块特例代码

### 开发环境

```powershell
dotnet build .\MediaFactory.slnx
dotnet run --project .\MediaFactory.App\MediaFactory.App.vbproj
```

### Pull Request 检查项

- 本地构建通过
- 没有把敏感信息加入受控文件
- 如果行为有变化，README 或配置模板也同步更新
- 如果涉及界面文案，已检查中英文两种语言下的效果
