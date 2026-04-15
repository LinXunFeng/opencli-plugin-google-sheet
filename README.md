# opencli-plugin-google-sheet

OpenCLI plugin for reading Google Sheets by `docId`.

> Language: English | [中文](https://github.com/LinXunFeng/opencli-plugin-google-sheet/blob/main/README-zh.md)

## ☕ Support me

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/T6T4JKVRP) [![wechat](https://img.shields.io/static/v1?label=WeChat&message=WeChat&nbsp;Pay&color=brightgreen&style=for-the-badge&logo=WeChat)](https://cdn.jsdelivr.net/gh/FullStackAction/PicBed@resource20220417121922/image/202303181116760.jpeg)

WeChat tech group: [Group details](https://mp.weixin.qq.com/s/JBbMstn0qW6M71hh-BRKzw)

## 📦 Install

```bash
# Local path install (replace with your absolute path)
opencli plugin install /absolute/path/to/opencli-plugin-google-sheet

# Update after install (common for local symlink setup)
opencli plugin update google-sheet
```

## 🧩 Commands

| Command | Description |
|---------|-------------|
| `opencli google-sheet sheets --docId <id>` | List worksheet tabs (`gid`, `title`, `index`) |
| `opencli google-sheet read --docId <id> --sheet <name\|gid>` | Read one worksheet as display values |

Notes:
- `read` defaults to JSON output (`--format json`).
- `read` also supports `table/csv/md/plain` (tabular output).
- If `--sheet` is omitted, `read` returns available worksheets first.

## ❓ FAQ

### 1) daemon / extension not connected

```bash
opencli doctor
opencli daemon stop
```

Then make sure Chrome/Chromium is open and the OpenCLI extension is enabled, then retry.

### 2) `sheets` returns only one result

First, verify whether `htmlview` contains multiple gids:

```bash
opencli browser open "https://docs.google.com/spreadsheets/d/<docId>/htmlview?rm=minimal"
opencli browser get html | rg -o 'gid=[0-9]+' | sort -u
```

If you can see multiple gids, read directly by gid:

```bash
opencli google-sheet read --docId <docId> --sheet <gid> -f table -v
```

### 3) `AUTH_REQUIRED` / `ACCESS_DENIED`

- Confirm the sheet can be opened in the same browser;
- Keep the account logged in;
- Retry the command.

## 🛠️ Development

```bash
# Enter plugin directory
cd /absolute/path/to/google-sheet

# Install locally (symlinked)
opencli plugin install "$PWD"

# Verify commands are registered successfully
opencli list | grep "google-sheet"

# Run tests
npm test
```

## 🖨 About Me

- GitHub: [https://github.com/LinXunFeng](https://github.com/LinXunFeng)
- Email: [linxunfeng@yeah.net](mailto:linxunfeng@yeah.net)
- Blogs:
  - FullStackAction: [https://fullstackaction.com](https://fullstackaction.com)
  - Juejin: [https://juejin.cn/user/1820446984512392](https://juejin.cn/user/1820446984512392)

<img height="267.5" width="481.5" src="https://github.com/LinXunFeng/LinXunFeng/raw/master/static/img/FSAQR.png"/>
