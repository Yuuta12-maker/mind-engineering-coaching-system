# マインドエンジニアリング・コーチング業務管理システム

Google Apps Script (GAS) を使用した、マインドエンジニアリング・コーチング（MEC）の業務管理システムです。

## システムの目的

マインドエンジニアリング・コーチングの業務効率化と顧客管理を強化するためのシステムです。
コーチング業務に特化した機能を提供し、クライアントとのスムーズなコミュニケーションを実現します。

## 主な機能

* クライアント情報管理
* セッション日程の管理と Google Calendar 連携
* 支払い管理
* 自動メール送信（リマインダー、次回セッション日程調整など）
* 契約書管理
* 領収書発行
* ダッシュボード
* 各種レポート生成

## 技術スタック

* Google Apps Script (GAS)
* Google Sheets（データベース）
* Google Forms（申込フォーム）
* Google Calendar（スケジュール管理）
* Google Docs（契約書テンプレート）
* HTML/CSS/JavaScript（Web UI）

## ディレクトリ構成

```
/
├── spreadsheet-setup/   # スプレッドシート初期化スクリプト
├── backend/             # GASバックエンドコード
│   ├── client/          # クライアント管理モジュール
│   ├── session/         # セッション管理モジュール
│   ├── payment/         # 支払い管理モジュール
│   ├── email/           # メール自動化モジュール
│   ├── document/        # 文書管理モジュール
│   └── utils/           # ユーティリティ関数
├── frontend/            # Web UI用コード
│   ├── dashboard/       # ダッシュボード
│   ├── client-management/  # クライアント管理画面
│   ├── session-management/ # セッション管理画面
│   ├── payment-management/ # 支払い管理画面
│   └── css/             # スタイルシート
└── docs/                # ドキュメント
```

## デプロイ方法

1. スプレッドシートを作成
2. スプレッドシート初期化スクリプトを実行
3. GASプロジェクトにコードをデプロイ
4. Web UIをデプロイ
5. 初期設定を行う

## 開発者

森山雄太
