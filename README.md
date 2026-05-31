# graphtool

Microsoft 365 の組織情報を Microsoft Graph API 経由で取得する CLI ツール。

## 前提条件

- `msal`, `requests` パッケージ
- Microsoft 365 アカウント（初回実行時にデバイスコード認証）

```bash
pip install msal requests
```

## 認証

初回実行時、またはトークンの期限切れ時にデバイスコード認証が発生します。  
表示された URL にアクセスしてコードを入力してください。トークンは `token_cache.json` にキャッシュされます。

```
% python graphtool.py login
To sign in, use a web browser to open the page https://login.microsoft.com/device and enter the code OW6QPKSPZ to authenticate.
```

ブラウザから URLを開いて、code (この例では OW6QPKSPZ) を入力します。  
認証に成功すると、コンソールに以下のようなメッセージが表示されます。

```
Token cached successfully.
```


## コマンド一覧

### `self` — 自分のユーザー情報を表示

```bash
python graphtool.py self
python graphtool.py self <user_id|UPN>
python graphtool.py self --format json
python graphtool.py self --format json --full
python graphtool.py self --save
```

### `manager` — マネージャーを表示

```bash
python graphtool.py manager
python graphtool.py manager <user_id|UPN>
python graphtool.py manager --format json
```

### `tree` — 組織ツリーを表示

指定ユーザーを起点に、直属の部下を再帰的に取得してツリー表示します。

```bash
python graphtool.py tree
python graphtool.py tree <user_id|UPN>
python graphtool.py tree --format json
python graphtool.py tree --format csv
python graphtool.py tree --format csv > org_tree.csv
python graphtool.py tree --save
```

CSV 出力のカラム: `displayName`, `userPrincipalName`, `jobTitle`, `managerUPN`（`managerUPN` が空の行がツリーのルート）

### `search` — ユーザーを検索

displayName・mail・UPN の部分一致で検索します（最大25件）。

```bash
python graphtool.py search "John"
python graphtool.py search "john@example.com"
python graphtool.py search "John" --format json
python graphtool.py search "John" --format json --full
```

### `search-groups` — グループを検索

displayName・mail の部分一致で検索します（最大25件）。

```bash
python graphtool.py search-groups "Engineering"
python graphtool.py search-groups "eng@example.com"
python graphtool.py search-groups "Engineering" --format json
python graphtool.py search-groups "Engineering" --format json --full
```

出力例（text）:
```
Engineering Team <eng-team@example.com>
  Description: Global engineering group
  Type: Microsoft 365
```

### `group-members` — グループメンバーを一覧表示

displayName またはメールアドレスでグループを指定します。

```bash
python graphtool.py group-members "Engineering Team"
python graphtool.py group-members eng-team@example.com
python graphtool.py group-members "Engineering Team" --format json
python graphtool.py group-members "Engineering Team" --format csv
python graphtool.py group-members "Engineering Team" --expand
python graphtool.py group-members "Engineering Team" --expand --format csv > members.csv
```

| オプション | 説明 |
|---|---|
| `--format text` | テキスト表示（デフォルト） |
| `--format json` | JSON 出力 |
| `--format csv` | CSV 出力（displayName, mail, jobTitle, userPrincipalName） |
| `--full` | JSON 出力時に全フィールドを含める |
| `--expand` | メンバーにグループが含まれる場合、再帰的に展開してユーザーのみ返す |

## 共通オプション

| オプション | 説明 |
|---|---|
| `--format text\|json` | 出力形式（`self`, `manager`, `search`, `search-groups`） |
| `--format text\|json\|csv` | 出力形式（`tree`, `group-members`） |
| `--full` | JSON 出力時に Graph API の全フィールドを含める |
| `--save` | 結果を JSON ファイルに保存（`self`, `manager`, `tree`） |

`user_id` を省略すると自分自身（`me`）が対象になります。UPN（例: `user@example.com`）またはオブジェクト ID を指定できます。

## 必要なスコープ

| コマンド | 必要なスコープ |
|---|---|
| `self`, `manager`, `tree`, `search` | `User.Read.All` |
| `search-groups`, `group-members` | `Group.Read.All` |

---

## MCP サーバー（graphtool_mcp.py）

`graphtool_mcp.py` は同じ機能を [Model Context Protocol (MCP)](https://modelcontextprotocol.io) サーバーとして提供します。Claude などの MCP クライアントから直接呼び出せます。

### 追加パッケージ

```bash
pip install mcp
```

### 提供ツール

| ツール | 引数 | 説明 |
|---|---|---|
| `get_user` | `user_id` (省略時 `"me"`) | ユーザー情報を取得 |
| `get_manager` | `user_id` (省略時 `"me"`) | マネージャーを取得 |
| `get_org_tree` | `user_id` (省略時 `"me"`) | 組織ツリーを JSON で取得 |
| `search_users` | `query` | 名前・メール・UPN の部分一致でユーザー検索 |
| `search_groups` | `query` | 名前・メールの部分一致でグループ検索 |
| `get_group_members` | `group_ref`, `expand` (省略時 `false`) | グループメンバーを CSV で取得。`expand=true` でネストグループを再帰展開 |

`user_id` には UPN（`user@example.com`）・オブジェクト ID・`"me"` を指定できます。  
`group_ref` にはグループの displayName またはメールアドレスを指定できます。

### Claude Code への登録

`~/.claude/claude_desktop_config.json` に追加します。

```json
{
  "mcpServers": {
    "graphtool": {
      "command": "/Users/foobar/envs/py3.12/bin/python",
      "args": ["/Users/foobar/opt/grahtool/graphtool_mcp.py"]
    }
  }
}
```

### Claude Code への登録後の確認

```bash
# 登録済みサーバーの確認（Claude Code TUI 内で）
/mcp
```

### Codex CLI への登録

`~/.codex/config.toml` に追加します（プロジェクトスコープの場合は `.codex/config.toml`）。

```toml
[mcp_servers.graphtool]
command = "/Users/foobar/envs/py3.12/bin/python"
args = ["/Users/foobar/opt/graphtool/graphtool_mcp.py"]
```

CLI コマンドで追加することもできます。

```bash
codex mcp add graphtool -- /Users/foobar/envs/py3.12/bin/python /Users/foobar/opt/graphtool/graphtool_mcp.py
```

登録済みサーバーは Codex TUI 内で `/mcp` コマンドで確認できます。

### 認証について

MCP サーバーも CLI と同じ `token_cache.json` を使用します。  
初回はデバイスコード認証が必要なため、先に CLI で認証を済ませておくことを推奨します。

```bash
python graphtool.py login
```
