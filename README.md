# Microsoft Graph Change Notifications with Delegated Permission

Microsoft Graph API の Change Notifications 機能と委任アクセス許可(Delegated Permission)を検証するためのプロジェクトです。ユーザーの受信トレイにメールが届いた際にリアルタイムで通知を受け取ることができます。

## 概要

このアプリケーションは以下の機能を提供します:

- **OAuth 2.0 認証**: MSAL (Microsoft Authentication Library) を使用したユーザー認証
- **Change Notifications サブスクリプション**: ユーザーの受信トレイの変更を監視
- **自動トークン更新**: アクセストークンの有効期限が切れる前に自動更新
- **自動サブスクリプション更新**: サブスクリプションの有効期限が切れる前に自動延長
- **Web UI**: ステータス確認用のシンプルな管理画面

## 前提条件

- Node.js 18.x 以降
- Azure AD に登録されたアプリケーション
- ngrok などのトンネリングツール（通知エンドポイントを公開するため）
- Microsoft 365 アカウント

## Azure AD アプリケーションの設定

1. [Azure Portal](https://portal.azure.com) にアクセス
2. Azure Active Directory > アプリの登録 > 新規登録
3. 以下の設定を行う:
   - **名前**: 任意の名前（例: Graph Subscription Demo）
   - **サポートされているアカウントの種類**: 任意の組織ディレクトリ内のアカウント
   - **リダイレクト URI**: Web - `http://localhost:3000/auth/callback`

4. アプリケーション登録後:
   - **概要** から `アプリケーション (クライアント) ID` と `ディレクトリ (テナント) ID` をコピー
   - **証明書とシークレット** > **新しいクライアント シークレット** を作成し、値をコピー

5. **API のアクセス許可** を設定:
   - Microsoft Graph > 委任されたアクセス許可 を追加
   - 以下の権限を追加:
     - `User.Read` - ユーザーのプロフィール読み取り
     - `Mail.Read` - メールの読み取り
     - `offline_access` - リフレッシュトークンの取得
   - **管理者の同意を与える** をクリック

## セットアップ

### 1. 依存関係のインストール

```bash
npm install
# または
pnpm install
```

### 2. 環境変数の設定

`.env.example` をコピーして `.env` ファイルを作成:

```bash
cp .env.example .env
```

`.env` ファイルを編集して、以下の値を設定:

```env
# Azure AD アプリケーション設定
CLIENT_ID=your-client-id-here
TENANT_ID=your-tenant-id-here
CLIENT_SECRET=your-client-secret-here

# OAuth リダイレクトURI
REDIRECT_URI=http://localhost:3000/auth/callback

# 通知エンドポイント (ngrok等の公開URL)
NOTIFICATION_URL=https://your-subdomain.ngrok.io/api/notifications

# サーバーポート
PORT=3000

# セキュリティ - クライアント状態検証用の秘密鍵
CLIENT_STATE_SECRET=your-random-secret-here
```

**CLIENT_STATE_SECRET の生成方法**:
```bash
openssl rand -base64 32
```

### 3. ngrok のセットアップ

Microsoft Graph が通知を送信するために、ローカルサーバーを公開する必要があります:

```bash
ngrok http 3000
```

ngrok が表示する HTTPS URL（例: `https://abc123.ngrok.io`）を `.env` の `NOTIFICATION_URL` に設定:

```env
NOTIFICATION_URL=https://abc123.ngrok.io/api/notifications
```

## 使い方

### 開発モードで起動

```bash
npm run dev
```

### 本番ビルドと起動

```bash
npm run build
npm start
```

## アプリケーションの使用手順

1. **サーバーを起動**
   ```bash
   npm run dev
   ```

2. **ブラウザでアクセス**

   ブラウザで `http://localhost:3000` を開く

3. **サインイン**

   「サインインしてサブスクライブ」ボタンをクリックして Microsoft アカウントでログイン

4. **認証完了**

   認証が成功すると、自動的に受信トレイの Change Notifications サブスクリプションが作成されます

5. **通知を受信**

   自分宛てにメールを送信すると、サーバーのコンソールに通知が表示されます:
   ```
   📬 通知を受信: 2024-01-10T12:34:56.789Z
   {
     "changeType": "created",
     "resource": "...",
     "clientState": "..."
   }
   ```

6. **ステータス確認**

   ブラウザで `http://localhost:3000/status` にアクセスすると、現在のトークンとサブスクリプションの状態を確認できます

## プロジェクト構造

```
.
├── src/
│   ├── app.ts              # メインアプリケーション (Express サーバー)
│   ├── auth.ts             # MSAL 認証ロジック
│   ├── graphClient.ts      # Microsoft Graph API クライアント
│   ├── renewalService.ts   # トークン/サブスクリプション自動更新サービス
│   ├── db.ts               # インメモリデータストア
│   ├── types/
│   │   └── index.ts        # TypeScript 型定義
│   └── config/
│       └── axios.ts        # Axios 設定
├── .env.example            # 環境変数のサンプル
├── package.json
├── tsconfig.json
└── README.md
```

## 主要なエンドポイント

| エンドポイント | メソッド | 説明 |
|--------------|---------|------|
| `/` | GET | ホームページ |
| `/auth/signin` | GET | OAuth 認証を開始 |
| `/auth/callback` | GET | OAuth コールバック |
| `/api/notifications` | POST | Graph からの Change Notifications を受信 |
| `/status` | GET | 現在のトークンとサブスクリプションの状態を表示 |
| `/health` | GET | ヘルスチェック API |

## 自動更新の仕組み

### トークンの自動更新

- 更新サービスが5分ごとに実行され、トークンの有効期限をチェック
- 有効期限の **10分前** になると自動的に更新
- MSAL のサイレントフロー (`acquireTokenSilent`) を使用

### サブスクリプションの自動更新

- 更新サービスが5分ごとに実行され、サブスクリプションの有効期限をチェック
- 有効期限の **1時間前** になると自動的に延長
- Microsoft Graph の PATCH `/subscriptions/{id}` API を使用
- 最大有効期限: 4230分（約3日間）

## トラブルシューティング

### 通知が届かない

1. **ngrok が起動しているか確認**
   ```bash
   ngrok http 3000
   ```

2. **NOTIFICATION_URL が正しいか確認**

   `.env` の `NOTIFICATION_URL` が ngrok の HTTPS URL と一致しているか確認

3. **サブスクリプションが有効か確認**

   `http://localhost:3000/status` でサブスクリプションの有効期限を確認

### 認証エラー

1. **Azure AD の設定を確認**
   - CLIENT_ID, TENANT_ID, CLIENT_SECRET が正しいか
   - リダイレクト URI が `http://localhost:3000/auth/callback` に設定されているか
   - API のアクセス許可が正しく設定され、管理者の同意が与えられているか

2. **ログを確認**

   サーバーのコンソール出力でエラーメッセージを確認

### サブスクリプション作成エラー

1. **NOTIFICATION_URL が HTTPS であることを確認**

   Microsoft Graph は HTTP の通知エンドポイントを受け付けません

2. **CLIENT_STATE_SECRET が設定されているか確認**

   `.env` に CLIENT_STATE_SECRET が設定されていることを確認

## 技術スタック

- **Node.js** / **TypeScript** - ランタイムと開発言語
- **Express** - Web フレームワーク
- **MSAL Node** - Microsoft 認証ライブラリ
- **Microsoft Graph Client** - Graph API SDK
- **Axios** - HTTP クライアント
- **dotenv** - 環境変数管理

## 制限事項

- **データストア**: 現在はインメモリストレージを使用。サーバー再起動でデータが消失します
- **本番環境**: 本番環境では PostgreSQL や MongoDB などの永続ストレージを使用してください
- **スケーラビリティ**: 単一インスタンスでの動作を想定。複数インスタンスで動作させる場合は共有ストレージが必要です
- **サブスクリプション有効期限**: メールボックスリソースの最大有効期限は4230分（約3日）です

## 参考リンク

- [Microsoft Graph Change Notifications](https://learn.microsoft.com/ja-jp/graph/api/resources/webhooks)
- [Microsoft Graph Subscriptions API](https://learn.microsoft.com/ja-jp/graph/api/subscription-post-subscriptions)
- [MSAL Node Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)
- [Microsoft Graph API Documentation](https://learn.microsoft.com/ja-jp/graph/overview)

## ライセンス

このプロジェクトは検証用のサンプルコードです。
