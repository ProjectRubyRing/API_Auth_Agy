import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_workbook():
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    # Common Font
    font_name = "Meiryo UI"
    
    # Monotone Color Palette
    COLOR_PRIMARY = "2F3542"      # Very Dark Slate / Black
    COLOR_SECONDARY = "57606F"    # Medium Dark Slate
    COLOR_HEADER_BG = "333333"    # Dark Charcoal Header
    COLOR_SUBHEADER_BG = "555555" # Medium Gray Subheader
    COLOR_TH_BG = "E4E7EB"        # Soft Gray Table Header
    COLOR_TH_DARK_BG = "2F3542"   # Dark Table Header
    COLOR_ZEBRA = "F8F9FA"        # Very Light Gray Zebra Row
    COLOR_WHITE = "FFFFFF"        # White
    COLOR_ACCENT_GRAY = "747D8C"  # Border/Accent Gray
    COLOR_NOTE_BG = "F1F2F6"      # Callout Box Light Gray
    COLOR_ALERT_BG = "E2E5E9"     # Slightly deeper gray for warnings/alerts

    font_title = Font(name=font_name, size=16, bold=True, color=COLOR_WHITE)
    font_sec_title = Font(name=font_name, size=12, bold=True, color=COLOR_WHITE)
    font_th = Font(name=font_name, size=10, bold=True, color=COLOR_WHITE)
    font_th_dark = Font(name=font_name, size=10, bold=True, color=COLOR_WHITE)
    font_th_light = Font(name=font_name, size=10, bold=True, color=COLOR_PRIMARY)
    font_body = Font(name=font_name, size=9.5, bold=False, color="1E272E")
    font_body_bold = Font(name=font_name, size=9.5, bold=True, color="1E272E")
    font_code = Font(name="Consolas", size=9, bold=False, color="2F3542")
    font_note = Font(name=font_name, size=9, italic=True, color="57606F")

    fill_title = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    fill_sec = PatternFill(start_color=COLOR_SUBHEADER_BG, end_color=COLOR_SUBHEADER_BG, fill_type="solid")
    fill_th_dark = PatternFill(start_color=COLOR_PRIMARY, end_color=COLOR_PRIMARY, fill_type="solid")
    fill_th_light = PatternFill(start_color=COLOR_TH_BG, end_color=COLOR_TH_BG, fill_type="solid")
    fill_zebra = PatternFill(start_color=COLOR_ZEBRA, end_color=COLOR_ZEBRA, fill_type="solid")
    fill_white = PatternFill(start_color=COLOR_WHITE, end_color=COLOR_WHITE, fill_type="solid")
    fill_note = PatternFill(start_color=COLOR_NOTE_BG, end_color=COLOR_NOTE_BG, fill_type="solid")
    fill_alert = PatternFill(start_color=COLOR_ALERT_BG, end_color=COLOR_ALERT_BG, fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CED6E0")
    medium_border_side = Side(border_style="medium", color="747D8C")
    double_border_side = Side(border_style="double", color="2F3542")

    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_th = Border(left=thin_border_side, right=thin_border_side, top=medium_border_side, bottom=medium_border_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    align_top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # -------------------------------------------------------------
    # SHEET 1: 表紙・全体概要 (Overview)
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="表紙・全体概要")
    ws1.views.sheetView[0].showGridLines = True

    # Title Banner
    ws1.merge_cells("B2:J3")
    cell_b2 = ws1["B2"]
    cell_b2.value = "API認証・認可アーキテクチャ徹底解説仕様書\n(イントラネット / Salesforce / インターネット / AWS Cognito M2M)"
    cell_b2.font = font_title
    cell_b2.fill = fill_title
    cell_b2.alignment = align_center

    # Meta Info
    meta_info = [
        ("作成目的", "API認証方式（API-Key / OAuth 2.0 / M2M）、Salesforce/インターネットAPIのトークン発行者、再利用設計、Cognito設定項目の完全解明"),
        ("対象読者", "システムアーキテクト、バックエンドエンジニア、性能試験担当者、セキュリティ担当者、初学者"),
        ("対象認証技術", "API-Key認証 (イントラネット), OAuth 2.0 Client Credentials Grant (M2M), AWS Cognito User Pool, Salesforce Connected App"),
        ("最新基準準拠", "RFC 6749 (OAuth 2.0), RFC 7519 (JWT), RFC 7009 (Token Revocation), AWS Cognito最新公式仕様"),
        ("ドキュメント構成", "5シート構成（本シート / API認証方式の比較 / トークンの再利用とTTL / Cognito M2M設定解説 / 小学生向け解説集）")
    ]
    row_idx = 5
    for label, val in meta_info:
        ws1.merge_cells(f"B{row_idx}:C{row_idx}")
        c_lbl = ws1[f"B{row_idx}"]
        c_lbl.value = label
        c_lbl.font = font_body_bold
        c_lbl.fill = fill_th_light
        c_lbl.alignment = align_center
        c_lbl.border = border_cell

        ws1.merge_cells(f"D{row_idx}:J{row_idx}")
        c_val = ws1[f"D{row_idx}"]
        c_val.value = val
        c_val.font = font_body
        c_val.fill = fill_white
        c_val.alignment = align_left
        c_val.border = border_cell
        row_idx += 1

    # Table of Contents
    row_idx += 2
    ws1.merge_cells(f"B{row_idx}:J{row_idx}")
    toc_head = ws1[f"B{row_idx}"]
    toc_head.value = "ドキュメント構成・目次（各シートの役割とハイライト）"
    toc_head.font = font_sec_title
    toc_head.fill = fill_sec
    toc_head.alignment = align_left
    row_idx += 1

    toc_items = [
        ("No.", "シート名", "主な記載内容", "得られる技術的インサイト・核心の結論"),
        ("1", "表紙・全体概要", "本仕様書の概要、前提技術、更新情報", "ドキュメント全体のスコープとエグゼクティブサマリ"),
        ("2", "API認証方式の比較", "イントラネットAPI (API-Key) / Salesforce API / インターネットAPI の認証・トークン発行者の違い", "Salesforce APIのトークン発行者はSalesforce自身。インターネットAPIの認可APIは実在しCognitoが発行。パフォーマンステスト影響評価。"),
        ("3", "トークン再利用とTTL", "アクセストークンの再利用可否、有効期間（TTL）の技術的定義、歴史的背景、キャッシュ設計", "トークン再利用は必須ベストプラクティス。都度取得によるCognito枯渇障害防止。有効期間のセキュリティと運用のトレードオフ。"),
        ("4", "Cognito_M2M設定解説", "Cognitoユーザプール各設定項目（ID/Secret/RefreshToken30日/AccessToken60分/IDToken1時間/Session3分/Revocation/UserExistence）の解剖", "M2M(Client Credentials)ではRefresh Token・ID Tokenはプロトコル上発行されない仕様の完全解説。各設定がM2Mに与える実際の影響を徹底整理。"),
        ("5", "小学生でもわかる解説集", "認証・認可の進化の歴史（Basic認証→API-Key→OAuth→JWT）、小学生でも直感的に理解できる例え話", "「合言葉」「テーマパークのパスポート」「入城許可証」に例えて誰でも直感的に本質を完全理解できるストーリー解説。")
    ]

    for i, item in enumerate(toc_items):
        is_h = (i == 0)
        ws1[f"B{row_idx}"] = item[0]
        ws1[f"C{row_idx}"] = item[1]
        ws1[f"D{row_idx}"] = item[2]
        ws1.merge_cells(f"E{row_idx}:J{row_idx}")
        ws1[f"E{row_idx}"] = item[3]

        for col_l in ["B", "C", "D", "E", "F", "G", "H", "I", "J"]:
            c = ws1[f"{col_l}{row_idx}"]
            c.font = font_th if is_h else font_body
            c.fill = fill_th_dark if is_h else (fill_zebra if i % 2 == 1 else fill_white)
            c.border = border_th if is_h else border_cell
            c.alignment = align_center if (is_h or col_l in ["B", "C"]) else align_left
        row_idx += 1

    # Key Summary Box
    row_idx += 2
    ws1.merge_cells(f"B{row_idx}:J{row_idx}")
    box_head = ws1[f"B{row_idx}"]
    box_head.value = "【エグゼクティブサマリ】ご質問に対する回答の結論一覧"
    box_head.font = font_sec_title
    box_head.fill = fill_sec
    box_head.alignment = align_left
    row_idx += 1

    summary_qas = [
        ("Q1. イントラネットAPIのAPI-Key認証", "閉域網内で高速・シンプルに通信するためHeaderにAPI-Keyを付与する方式。追加の認可サーバー通信が不要で軽量。"),
        ("Q2. Salesforce APIの認証とトークン発行者", "OAuth 2.0認証を採用。アクセストークンの発行者は「Salesforce自身の認可エンドポイント（/services/oauth2/token）」。"),
        ("Q3. インターネットAPIの認可APIの存在と発行者", "OAuth 2.0認証であり、認可API（トークン取得エンドポイント）は「実際に必ず存在している」。トークン発行者は「AWS Cognito ユーザプール」。仕様書で業務APIと認証APIが別紙分割されているため未記載に見えるだけ。"),
        ("Q4. パフォーマンステストでの影響度", "トークンを適切にキャッシュ・再利用（60分TTLで55分間再利用等）すれば、トークン発行回数は1時間に1回程度となり「無視できる水準」。ただし、キャッシュせずリクエスト毎に都度取得するとCognitoのスロットリングを誘発しシステム全体が停止する重大リスクあり。"),
        ("Q5. アクセストークンの再利用と有効期間", "再利用は「可能であり、強く推奨（必須のベストプラクティス）」。有効期間（TTL）とは「万が一トークンが漏洩しても自動的に無効化される安全限界時間」。"),
        ("Q6. Cognito M2M設定の本質", "M2M(Client Credentials)では「Access Token(60分)」と「Client ID/Secret」のみが実働する。「Refresh Token(30日)」と「ID Token(1時間)」はOAuth 2.0仕様上M2Mでは発行されない。設定画面は人間向け認証と共通のため混在表示されている。")
    ]

    for q, a in summary_qas:
        ws1.merge_cells(f"B{row_idx}:D{row_idx}")
        ws1[f"B{row_idx}"] = q
        ws1[f"B{row_idx}"].font = font_body_bold
        ws1[f"B{row_idx}"].fill = fill_th_light
        ws1[f"B{row_idx}"].border = border_cell
        ws1[f"B{row_idx}"].alignment = align_left

        ws1.merge_cells(f"E{row_idx}:J{row_idx}")
        ws1[f"E{row_idx}"] = a
        ws1[f"E{row_idx}"].font = font_body
        ws1[f"E{row_idx}"].fill = fill_white
        ws1[f"E{row_idx}"].border = border_cell
        ws1[f"E{row_idx}"].alignment = align_left
        row_idx += 1

    # -------------------------------------------------------------
    # SHEET 2: API認証方式の比較 (Intranet vs Salesforce vs Internet)
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="API認証方式の比較")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H2")
    ws2["A1"] = "API認証方式の徹底比較：イントラネットAPI / Salesforce API / インターネットAPI"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_title
    ws2["A1"].alignment = align_center

    # Detailed Comparison Table
    headers_s2 = ["比較項目", "イントラネットAPI", "Salesforce API (外部SaaS)", "インターネットAPI (自社公開API)"]
    ws2.merge_cells("A4:B4")
    ws2["A4"] = headers_s2[0]
    ws2.merge_cells("C4:D4")
    ws2["C4"] = headers_s2[1]
    ws2.merge_cells("E4:F4")
    ws2["E4"] = headers_s2[2]
    ws2.merge_cells("G4:H4")
    ws2["G4"] = headers_s2[3]

    for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        c = ws2[f"{col}4"]
        c.font = font_th
        c.fill = fill_th_dark
        c.border = border_th
        c.alignment = align_center

    comp_data = [
        ("認証・認可方式", 
         "API-Key認証 (固定共有鍵方式)", 
         "OAuth 2.0 (Client CredentialsまたはJWT Bearer Flow)", 
         "OAuth 2.0 (Client Credentials Grant / M2M)"),
        ("トークンの形態", 
         "静的な固定文字列 (例: x-api-key: secret_abc123)", 
         "動的アクセストークン (SalesforceセッションID / Bearer Token)", 
         "動的アクセストークン (署名付きステートレスJWT)"),
        ("アクセストークンの発行者は誰か？", 
         "発行という概念なし（事前払出の固定キーを静的比較）", 
         "【Salesforce自身】\nSalesforceログインエンドポイント(/services/oauth2/token)", 
         "【AWS Cognito ユーザプール】\nCognitoのOAuthトークンエンドポイント(/oauth2/token)"),
        ("API認可トークン取得APIの存在有無", 
         "存在しない（キーは事前配布され、ヘッダに直設定）", 
         "存在する (SalesforceのOAuth 2.0 Token API)", 
         "【必ず存在する】\n業務API仕様書とは別紙（共通認証仕様）として管理されているため未記載に見える"),
        ("動作原理と通信の流れ", 
         "① クライアントがリクエストヘッダにAPI-Keyを付与\n② API Gateway/内部アプリがマスターキーと完全一致検証\n③ 一致すればAPIを実行", 
         "① クライアントがSalesforceへクライアント証明書/鍵でトークン要求\n② Salesforceが認証しアクセストークンを返却\n③ トークンをAuthorizationヘッダに付与してSalesforce REST API呼出", 
         "① クライアントがAWS CognitoへClient ID/Secretでトークン要求\n② Cognitoが署名付きJWTアクセストークンを発行\n③ クライアントがJWTをAuthorization: Bearerに付与してAPI呼出\n④ API GatewayがCognito公開鍵(JWKS)で署名検証"),
        ("主な採用理由・背景", 
         "社内閉域網(VPN/DirectConnect)内であるため傍受リスクが低く、通信レイテンシ極小化・実装単純化を最優先するため", 
         "外部SaaSとしての標準プロトコル。Salesforce内の権限(プロファイル/権限セット)と接続アプリを厳密に紐付けるため", 
         "インターネット経由の第三者アクセスに対し、固定パスワードを渡さず、漏洩しても短時間で失効する安全な仕組みが必須なため"),
        ("メリット", 
         "・実装が極めて容易（数行のコード）\n・認証サーバー不要で追加の通信レイテンシがゼロ\n・高スループット内部通信に最適", 
         "・Salesforceのきめ細かい権限管理と完全連動\n・クライアント毎のアクセス監視・即時停止が可能", 
         "・ステートレス検証によりAPIサーバー負荷が極小\n・有効期限(60分)による漏洩リスク極小化\n・OAuth 2.0標準仕様準拠"),
        ("デメリット・セキュリティ上の注意点", 
         "・キー漏洩時に無制限で悪用される\n・キーローテーション(変更)時に全システム一斉切替が必要\n・「誰が呼んだか」の監査性が低い", 
         "・Salesforce API呼出制限(ガバナ制限)を消費\n・Salesforce側のトークンエンドポイント障害に依存", 
         "・初回トークン取得のオーバーヘッド(TLS+署名)\n・トークンキャッシュ設計が必須（都度取得は厳禁）"),
        ("パフォーマンステスト観点の影響度", 
         "【大勢に影響なし】\n内部比較のみで処理負荷は数マイクロ秒レベル。ボトルネックになり得ない", 
         "【本番・外部依存】\nテスト時は通常モック化またはSalesforceサンドボックス呼出制限に留意が必要", 
         "【キャッシュ有無で激変】\n・キャッシュあり(1時間に1回取得): 大勢に影響なし(無視可能)\n・都度取得: Cognitoが即スロットリング(429)を起こし大惨事")
    ]

    r2 = 5
    for row in comp_data:
        ws2.merge_cells(f"A{r2}:B{r2}")
        ws2[f"A{r2}"] = row[0]
        ws2[f"A{r2}"].font = font_body_bold
        ws2[f"A{r2}"].fill = fill_th_light
        ws2[f"A{r2}"].alignment = align_top_left
        ws2[f"A{r2}"].border = border_cell

        ws2.merge_cells(f"C{r2}:D{r2}")
        ws2[f"C{r2}"] = row[1]
        ws2[f"C{r2}"].font = font_body
        ws2[f"C{r2}"].fill = fill_white
        ws2[f"C{r2}"].alignment = align_top_left
        ws2[f"C{r2}"].border = border_cell

        ws2.merge_cells(f"E{r2}:F{r2}")
        ws2[f"E{r2}"] = row[2]
        ws2[f"E{r2}"].font = font_body
        ws2[f"E{r2}"].fill = fill_zebra if (r2 % 2 == 1) else fill_white
        ws2[f"E{r2}"].alignment = align_top_left
        ws2[f"E{r2}"].border = border_cell

        ws2.merge_cells(f"G{r2}:H{r2}")
        ws2[f"G{r2}"] = row[3]
        ws2[f"G{r2}"].font = font_body
        ws2[f"G{r2}"].fill = fill_white
        ws2[f"G{r2}"].alignment = align_top_left
        ws2[f"G{r2}"].border = border_cell
        r2 += 1

    # Deep Dive: Why is Token API not documented on Internet API side?
    r2 += 2
    ws2.merge_cells(f"A{r2}:H{r2}")
    ws2[f"A{r2}"] = "【深掘り解説】インターネットAPI側に「API認可トークン取得API」の記載がない理由の完全究明"
    ws2[f"A{r2}"].font = font_sec_title
    ws2[f"A{r2}"].fill = fill_sec
    ws2[f"A{r2}"].alignment = align_left
    r2 += 1

    reasons = [
        ("理由①：関心の分離（アーキテクチャ上の標準設計）", 
         "Web APIの設計において、「業務API（リソースサーバー：/orders, /usersなど）」と「認証認可サーバー（IdP：Cognito /oauth2/token）」は物理的・論理的に完全に分かれています。SwaggerやOpenAPI仕様書は業務APIの入出力を定義する文書であるため、認可サーバーのエンドポイントは『共通接続マニュアル』『セキュリティ編』として別冊化されるのが業界標準です。"),
        ("理由②：OAuth 2.0 (RFC 6749) 標準プロトコルへの準拠", 
         "独自に作成したAPIではなく、AWS Cognitoのマネージドドメインが提供する標準仕様（https://<domain>.auth.<region>.amazoncognito.com/oauth2/token）をそのまま利用しているため、業務API開発チームがドキュメントを作成せず、AWS公式のトークン取得手順を参照する前提になっていることが多々あります。"),
        ("理由③：クライアント（呼び出し元）への事前設定提供", 
         "M2M接続を行う相手方（パートナー企業や自社他拠点）に対しては、API仕様書とは別に「Client ID」「Client Secret」「トークン発行URL」「スコープ」を暗号化経路で個別に払い出す運用が一般的です。そのため外部公開のAPI一覧には記載されません。"),
        ("結論（現場での確認・実態）", 
         "「記載がない＝存在しない」ではなく、「【必ずインターネット上にCognitoの認可エンドポイントが存在している】」というのが100%確実な技術的実態です。これがないとOAuth 2.0で保護されたインターネットAPIを呼び出すことが原理的に不可能です。")
    ]

    for title, desc in reasons:
        ws2.merge_cells(f"A{r2}:C{r2}")
        ws2[f"A{r2}"] = title
        ws2[f"A{r2}"].font = font_body_bold
        ws2[f"A{r2}"].fill = fill_th_light
        ws2[f"A{r2}"].alignment = align_top_left
        ws2[f"A{r2}"].border = border_cell

        ws2.merge_cells(f"D{r2}:H{r2}")
        ws2[f"D{r2}"] = desc
        ws2[f"D{r2}"].font = font_body
        ws2[f"D{r2}"].fill = fill_white
        ws2[f"D{r2}"].alignment = align_top_left
        ws2[f"D{r2}"].border = border_cell
        r2 += 1

    # Deep Dive: Performance Test Impact
    r2 += 2
    ws2.merge_cells(f"A{r2}:H{r2}")
    ws2[f"A{r2}"] = "【深掘り考察】パフォーマンステスト観点での「インターネットAPIは対象外なので無視できる水準」の妥当性と盲点"
    ws2[f"A{r2}"].font = font_sec_title
    ws2[f"A{r2}"].fill = fill_sec
    ws2[f"A{r2}"].alignment = align_left
    r2 += 1

    perf_points = [
        ("観点①：トークンキャッシュ前提なら「無視できる」は正しい", 
         "クライアント側がCognitoから取得したアクセストークン（有効期限60分）をメモリ等にキャッシュし、有効期限内（例: 55分間）再利用する設計になっている場合、Cognitoへのトークン取得リクエストは「システム全体で1時間に1回程度」となります。この場合、負荷試験のトラフィック全体から見て0.001%未満であり、性能影響は名実ともに『無視できる水準（大勢に影響なし）』と判断して差し支えありません。"),
        ("観点②：【潜む罠】もし「都度取得（キャッシュなし）」だった場合の大惨事", 
         "もしクライアントの実装が、APIを1回呼ぶたびに律儀にCognitoのトークン取得APIを呼ぶ実装（アンチパターン）になっていた場合、事態は一変します。\n1. CognitoのClient Credentialsエンドポイントには秒間リクエスト上限（ソフトリミット/クォータ）があり、高負荷時に即座にスロットリング（HTTP 429 Too Many Requests）が発生します。\n2. トークン発行にはTLSハンドシェイクと非対称鍵（RSA256）による電子署名生成が伴うため、API呼出自体の2倍〜3倍のレイテンシ遅延が加算されます。\n3. 結果として、パフォーマンステストでは合格したのに、本番リリース後に外部連携システムがCognitoをパンクさせ、内部APIまで巻き添えで障害停止するリスクがあります。"),
        ("観点③：性能試験担当者・アーキテクトが打つべき確認アクション", 
         "「インターネットAPIは対象外だから無視」で片付けるのではなく、呼び出し元システムの設計書またはソースコードを確認し、『【トークンがクライアント側で確実にキャッシュ・再利用されているか】』の1点だけは非機能要件として確認しておくことが極めて推奨されます。")
    ]

    for title, desc in perf_points:
        ws2.merge_cells(f"A{r2}:C{r2}")
        ws2[f"A{r2}"] = title
        ws2[f"A{r2}"].font = font_body_bold
        ws2[f"A{r2}"].fill = fill_alert
        ws2[f"A{r2}"].alignment = align_top_left
        ws2[f"A{r2}"].border = border_cell

        ws2.merge_cells(f"D{r2}:H{r2}")
        ws2[f"D{r2}"] = desc
        ws2[f"D{r2}"].font = font_body
        ws2[f"D{r2}"].fill = fill_white
        ws2[f"D{r2}"].alignment = align_top_left
        ws2[f"D{r2}"].border = border_cell
        r2 += 1

    # -------------------------------------------------------------
    # SHEET 3: トークンの再利用とTTL (Token Reuse & TTL)
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="トークンの再利用とTTL")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:H2")
    ws3["A1"] = "アクセストークンの再利用可能性と有効期間（TTL）の本質解説"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_title
    ws3["A1"].alignment = align_center

    # Section 1: Is Token Reusable?
    ws3.merge_cells("A4:H4")
    ws3["A4"] = "1. アクセストークンの再利用は可能か？（結論：可能であり、必須のベストプラクティス）"
    ws3["A4"].font = font_sec_title
    ws3["A4"].fill = fill_sec
    ws3["A4"].alignment = align_left

    reuse_points = [
        ("再利用の可否", "【完全に可能】発行されたアクセストークン（Bearer Token / JWT）は、その有効期間（60分）が終了するまでの間、何回でも・何百回でも再利用できます。"),
        ("なぜ再利用可能なのか（動作原理）", "アクセストークンはステートレスなJWT（JSON Web Token）です。トークン内部に有効期限(exp)や権限(scope)が暗号署名されており、API Gateway等の検証側は「Cognitoの公開鍵で署名が正しいか」「現在時刻 < expか」だけを数学的にチェックします。Cognitoに毎回問い合わせる必要がないため、同じトークンを何度使っても同じように認可が通ります。"),
        ("なぜ再利用しなければならないのか？", "① パフォーマンス：毎回トークンを取ると「トークン取得(約100〜200ms)」+「API呼出(約50ms)」となり、通信時間が3倍以上に跳ね上がります。\n② 認可サーバー保護：Cognitoのトークン発行エンドポイントに無駄な負荷をかけ、レートリミット超過（429エラー）でシステム全体が停止するのを防ぐためです。"),
        ("推奨される実装設計（キャッシュ戦略）", "クライアントはメモリ内にトークンと有効期限を保持します。\n【ベストプラクティス】：有効期限（60分）の5分前（55分経過時点）まではキャッシュされたトークンを使い回し、残り5分を切ったらバックグラウンドまたは次回API呼出時に新しいトークンを取得してキャッシュを上書きします（マージン設計）。")
    ]

    r3 = 5
    for lbl, val in reuse_points:
        ws3.merge_cells(f"A{r3}:B{r3}")
        ws3[f"A{r3}"] = lbl
        ws3[f"A{r3}"].font = font_body_bold
        ws3[f"A{r3}"].fill = fill_th_light
        ws3[f"A{r3}"].alignment = align_top_left
        ws3[f"A{r3}"].border = border_cell

        ws3.merge_cells(f"C{r3}:H{r3}")
        ws3[f"C{r3}"] = val
        ws3[f"C{r3}"].font = font_body
        ws3[f"C{r3}"].fill = fill_white
        ws3[f"C{r3}"].alignment = align_top_left
        ws3[f"C{r3}"].border = border_cell
        r3 += 1

    # Section 2: What is Expiration Time (TTL)?
    r3 += 2
    ws3.merge_cells(f"A{r3}:H{r3}")
    ws3[f"A{r3}"] = "2. 「有効期間（TTL: Time To Live / Expiration）」とは何か？ なぜ必要なのか？"
    ws3[f"A{r3}"].font = font_sec_title
    ws3[f"A{r3}"].fill = fill_sec
    ws3[f"A{r3}"].alignment = align_left
    r3 += 1

    ttl_points = [
        ("有効期間の定義", "トークンが発行された瞬間から、通行証として認められなくなる（無効化される）までの寿命時間のこと。Cognitoの初期値では「60分」。JWT内部の `exp` (Expiration Time) クレームにUNIXエポック秒として刻印されます。"),
        ("なぜ有効期間が必要なのか？（歴史的経緯）", "【無期限キーの悲劇からの教訓】\nかつてのAPIキーやパスワード認証は「有効期限がない（無期限）」のが当たり前でした。しかし、ソースコードへの誤コミット、通信の盗聴、ログファイルへの出力等によってキーが一度外部に漏洩すると、管理者が気づいて手動でキーを削除・変更するまで、攻撃者が何ヶ月もシステムに侵入し放題になるという致命的な事件が世界中で多発しました。"),
        ("ステートレスJWT特有のセキュリティジレンマ", "OAuth 2.0で採用されているJWTは「自己完結型（ステートレス）」です。サーバー側でセッションDBを持たずに高速検証できる反面、「一度発行したJWTを途中で強制無効化（ブラックリスト化）するのが技術的に極めて難しい」という弱点があります。\nそこで、「【最初から寿命を短く（例: 60分）設定しておき、万が一トークンを泥棒に盗まれても、60分経てば自動的に無価値なゴミ屑になる】」という時限爆弾式の安全装置として考案されたのが有効期間です。"),
        ("有効期間のトレードオフ", "・TTLが短すぎる場合（例: 1分）：セキュリティは最高だが、頻繁にCognitoへトークン再取得が発生し負荷・遅延が増大。\n・TTLが長すぎる場合（例: 24時間）：再取得回数は減るが、漏洩時に悪用される危険時間が長くなる。\n・バランス（標準値 60分）：M2M連携において、セキュリティとシステム負荷のバランスが最も取れた世界標準の業界デファクト設定。")
    ]

    for lbl, val in ttl_points:
        ws3.merge_cells(f"A{r3}:B{r3}")
        ws3[f"A{r3}"] = lbl
        ws3[f"A{r3}"].font = font_body_bold
        ws3[f"A{r3}"].fill = fill_th_light
        ws3[f"A{r3}"].alignment = align_top_left
        ws3[f"A{r3}"].border = border_cell

        ws3.merge_cells(f"C{r3}:H{r3}")
        ws3[f"C{r3}"] = val
        ws3[f"C{r3}"].font = font_body
        ws3[f"C{r3}"].fill = fill_white
        ws3[f"C{r3}"].alignment = align_top_left
        ws3[f"C{r3}"].border = border_cell
        r3 += 1

    # Section 3: Performance Comparison: Reusing vs Fetching Every Time
    r3 += 2
    ws3.merge_cells(f"A{r3}:H{r3}")
    ws3[f"A{r3}"] = "3. 性能検証：トークン都度取得 vs トークンキャッシュ再利用の定量比較（10,000リクエスト想定）"
    ws3[f"A{r3}"].font = font_sec_title
    ws3[f"A{r3}"].fill = fill_sec
    ws3[f"A{r3}"].alignment = align_left
    r3 += 1

    perf_headers = ["評価項目", "アンチパターン：リクエスト毎に都度取得", "ベストプラクティス：トークン再利用（60分キャッシュ）", "性能改善効果・技術的影響"]
    ws3.merge_cells(f"A{r3}:B{r3}")
    ws3[f"A{r3}"] = perf_headers[0]
    ws3.merge_cells(f"C{r3}:D{r3}")
    ws3[f"C{r3}"] = perf_headers[1]
    ws3.merge_cells(f"E{r3}:F{r3}")
    ws3[f"E{r3}"] = perf_headers[2]
    ws3.merge_cells(f"G{r3}:H{r3}")
    ws3[f"G{r3}"] = perf_headers[3]

    for col in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        c = ws3[f"{col}{r3}"]
        c.font = font_th
        c.fill = fill_th_dark
        c.border = border_th
        c.alignment = align_center
    r3 += 1

    perf_table = [
        ("Cognitoへのトークン取得回数", "10,000 回", "3 〜 4 回（約3時間のバッチ処理の場合）", "【99.96%削減】Cognitoへの負荷をほぼゼロ化"),
        ("通信オーバーヘッド (合計遅延)", "約 1,500秒 (10,000回 × 150ms)", "約 0.6秒 (4回 × 150ms)", "全体処理時間を大幅短縮（約25分間の無駄な待機時間を削減）"),
        ("1リクエストあたりの平均レスポンス", "約 200 ms (認証150ms + API 50ms)", "約 50 ms (キャッシュ取得0ms + API 50ms)", "【4倍高速化】クライアント体感速度が劇的に向上"),
        ("Cognitoスロットリング発生リスク", "【極めて危険】秒間数十リクエストで即429制限発生", "【皆無】クォータ制限に引っかかる可能性ゼロ", "本番環境での巻き添え障害を完全防止"),
        ("ネットワーク帯域・コスト消費", "毎リクエストでTLSネゴシエーションとJSON送受信が発生", "同一接続で実APIリクエストのみを送信", "不要なトラフィックとインフラコストの最小化")
    ]

    for r_item in perf_table:
        ws3.merge_cells(f"A{r3}:B{r3}")
        ws3[f"A{r3}"] = r_item[0]
        ws3[f"A{r3}"].font = font_body_bold
        ws3[f"A{r3}"].fill = fill_th_light
        ws3[f"A{r3}"].alignment = align_top_left
        ws3[f"A{r3}"].border = border_cell

        ws3.merge_cells(f"C{r3}:D{r3}")
        ws3[f"C{r3}"] = r_item[1]
        ws3[f"C{r3}"].font = font_body
        ws3[f"C{r3}"].fill = fill_alert
        ws3[f"C{r3}"].alignment = align_top_left
        ws3[f"C{r3}"].border = border_cell

        ws3.merge_cells(f"E{r3}:F{r3}")
        ws3[f"E{r3}"] = r_item[2]
        ws3[f"E{r3}"].font = font_body
        ws3[f"E{r3}"].fill = fill_white
        ws3[f"E{r3}"].alignment = align_top_left
        ws3[f"E{r3}"].border = border_cell

        ws3.merge_cells(f"G{r3}:H{r3}")
        ws3[f"G{r3}"] = r_item[3]
        ws3[f"G{r3}"].font = font_body_bold
        ws3[f"G{r3}"].fill = fill_zebra
        ws3[f"G{r3}"].alignment = align_top_left
        ws3[f"G{r3}"].border = border_cell
        r3 += 1

    # -------------------------------------------------------------
    # SHEET 4: Cognito_M2M設定徹底解剖 (Cognito M2M Settings)
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="Cognito_M2M設定解説")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:I2")
    ws4["A1"] = "AWS Cognito ユーザプール設定の徹底解剖（M2M認証前提における挙動と真実）"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_title
    ws4["A1"].alignment = align_center

    # Notice Box on M2M Protocol
    ws4.merge_cells("A4:I5")
    ws4["A4"] = "【最重要アーキテクチャ前提】：M2M（Machine-to-Machine）認証＝OAuth 2.0 Client Credentials Grant（RFC 6749 Section 4.4）\nAWS Cognitoの管理画面は「人間向け認証（ログイン画面/パスワード/MFA）」と「M2M認証（サーバー間連携）」で設定項目が一元化されているため、M2Mではプロトコル上使われない項目が存在します。各設定が『実際にバリバリ動くもの』か『存在はするがM2Mでは使われないもの』かを完全に切り分けて解説します。"
    ws4["A4"].font = font_body_bold
    ws4["A4"].fill = fill_note
    ws4["A4"].alignment = align_left
    ws4["A4"].border = border_cell

    # Settings Deep Dive Table
    r4 = 7
    s4_headers = ["設定項目名", "設定値", "M2M認証における実態・有効性", "プロトコルレベルの動作原理と動作イメージ", "歴史・背景と存在理由・推奨プラクティス"]
    ws4.merge_cells(f"A{r4}:B{r4}")
    ws4[f"A{r4}"] = s4_headers[0]
    ws4[f"C{r4}"] = s4_headers[1]
    ws4.merge_cells(f"D{r4}:E{r4}")
    ws4[f"D{r4}"] = s4_headers[2]
    ws4.merge_cells(f"F{r4}:G{r4}")
    ws4[f"F{r4}"] = s4_headers[3]
    ws4.merge_cells(f"H{r4}:I{r4}")
    ws4[f"H{r4}"] = s4_headers[4]

    for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
        c = ws4[f"{col}{r4}"]
        c.font = font_th
        c.fill = fill_th_dark
        c.border = border_th
        c.alignment = align_center
    r4 += 1

    settings_detail = [
        ("クライアントID (Client ID)",
         "英数字の一意識別子",
         "【必須・最重要】\n直接フル稼働",
         "M2M通信を行うシステム（クライアント）の「公開名義・ログインID」。トークン取得リクエストのボディまたはAuthorization: Basicヘッダに設定してCognitoに送信します。",
         "【背景】：どのシステムがアクセスしてきたかを特定し、許可されたスコープ(custom:api/read等)を識別するために必須。公開されても直ちに致命的ではないが、Secretとセットで厳重管理が基本。"),
        
        ("クライアントシークレット (Client Secret)",
         "暗号学的ランダム文字列",
         "【必須・最重要】\n直接フル稼働",
         "クライアントシステム自身の「パスワード・秘密鍵」。バックエンドサーバー同士の通信（Confidential Client）であるため、クライアント側に安全に保管されている前提で認証が行われます。",
         "【背景】：ブラウザやスマホアプリ（Public Client）ではシークレットを隠せませんが、サーバー（M2M）は環境変数等で安全に隠蔽可能。もしシークレット漏洩が疑われる場合はアプリクライアントの再作成またはシークレット再生成が必要。"),
        
        ("更新トークンの有効期限 (Refresh Token)",
         "30日",
         "【M2Mでは不使用・発行されない】\n※要注意ポイント！",
         "OAuth 2.0仕様（RFC 6749 Section 4.4.3）において「Client CredentialsフローではRefresh Tokenを含めてはならない（SHOULD NOT）」と規定されています。そのためCognitoのM2Mエンドポイントは、Refresh Tokenを一切発行しません。",
         "【なぜ30日と設定されているのか】：Cognitoユーザプールのアプリクライアント作成画面で「全認証フロー共通」のデフォルト値として30日が表示されているためです。M2Mではトークン切れ時にClient Secretを使って直接Access Tokenを再取得するため、Refresh Tokenの概念自体が不要です。"),
        
        ("アクセストークンの有効期限 (Access Token)",
         "60分",
         "【主役・直接作用】\n最も重要な稼働設定",
         "M2M通信でAPIを呼び出すためのBearerトークン（JWT）の有効期間。発行から60分間、API Gateway等でステートレスに検証・認可されます。JWTのexpクレームに現在時刻+3600秒が刻印されます。",
         "【背景】：業界標準値（デフォルト60分）。短すぎると頻繁にCognitoにトークンを取りに行くためCognitoスロットリングや性能遅延を招き、長すぎると万が一の漏洩時の悪用リスクが伸びます。クライアント側は「55分間キャッシュして再利用」するのが黄金律です。"),
        
        ("IDトークンの有効期限 (ID Token)",
         "1時間",
         "【M2Mでは不使用・発行されない】",
         "OpenID Connect (OIDC) 仕様上、IDトークンは「ログインした人間（ユーザー）の身元属性（メールアドレス、氏名、sub等）」をフロントエンドに伝えるためのものです。M2Mには人間が存在しないため、CognitoはIDトークンを発行しません。",
         "【なぜ1時間と設定されているのか】：こちらもアプリクライアント設定の共通項目として画面に表示されているためです。M2Mの /oauth2/token レスポンスには access_token, expires_in, token_type: 'Bearer' の3点のみが返却されます。"),
        
        ("認証フローセッションの持続期間",
         "3分",
         "【M2Mでは直接不使用】",
         "人間がCognitoの認証API（InitiateAuthなど）を呼び出し、パスワード入力後に「SMS認証コードやワンタイムパスワード（TOTP）を入力するまでの待ち時間（チャレンジセッション）」の有効期限（デフォルト3分）です。",
         "【背景】：M2MのClient Credentialsフローは、HTTP POSTの1リクエスト・1レスポンスで即座に完結（トークンが即返却）するため、途中のチャレンジ待ちセッションが発生しません。したがってM2M通信時には一切使われません。"),
        
        ("トークンの取消を有効化 (Token Revocation)",
         "有効化",
         "【M2Mでは実質影響限定的】\n※アーキテクチャ上の注意",
         "RFC 7009に基づくトークン失効API（/oauth2/revoke）を有効にする設定です。主に「人間がログアウトした際にリフレッシュトークンを即座に破棄する」ために利用されます。",
         "【M2Mでの注意点】：M2Mで発行されるアクセストークンはステートレスなJWTです。Cognito側でRevoke APIを叩いても、API Gateway側が公開鍵（JWKS）でオフライン検証している場合、API Gatewayは有効期限（60分）が切れるまでトークンが失効したことを検知できません。即時失効させたい場合はClient Secretをローテーションする必要があります。"),
        
        ("ユーザ存在エラーの防止 (Prevent User Existence)",
         "有効化",
         "【セキュリティベースライン】\n※直接の攻撃対象外だが推奨",
         "ユーザー認証時に存在しないユーザー名を指定した場合でも、UserNotFoundExceptionではなく常にNotAuthorizedException（ユーザー名またはパスワードが無効）を返し、攻撃者によるユーザーIDの有無調査（列挙攻撃）を防ぐ機能です。",
         "【M2Mでの位置づけ】：M2M認証ではユーザープール内の「ユーザー（エンドユーザー）」ではなく「アプリクライアント」を認証するため直接は関係しません。しかし、同一クライアントで将来的にユーザー認証を行う可能性や、プール全体のセキュリティ強化として有効化しておくのがAWSベストプラクティスです。")
    ]

    for item in settings_detail:
        ws4.merge_cells(f"A{r4}:B{r4}")
        ws4[f"A{r4}"] = item[0]
        ws4[f"A{r4}"].font = font_body_bold
        ws4[f"A{r4}"].fill = fill_th_light
        ws4[f"A{r4}"].alignment = align_top_left
        ws4[f"A{r4}"].border = border_cell

        ws4[f"C{r4}"] = item[1]
        ws4[f"C{r4}"].font = font_body_bold
        ws4[f"C{r4}"].fill = fill_white
        ws4[f"C{r4}"].alignment = align_top_center
        ws4[f"C{r4}"].border = border_cell

        ws4.merge_cells(f"D{r4}:E{r4}")
        ws4[f"D{r4}"] = item[2]
        ws4[f"D{r4}"].font = font_body_bold
        ws4[f"D{r4}"].fill = fill_zebra if "不使用" in item[2] else fill_white
        ws4[f"D{r4}"].alignment = align_top_left
        ws4[f"D{r4}"].border = border_cell

        ws4.merge_cells(f"F{r4}:G{r4}")
        ws4[f"F{r4}"] = item[3]
        ws4[f"F{r4}"].font = font_body
        ws4[f"F{r4}"].fill = fill_white
        ws4[f"F{r4}"].alignment = align_top_left
        ws4[f"F{r4}"].border = border_cell

        ws4.merge_cells(f"H{r4}:I{r4}")
        ws4[f"H{r4}"] = item[4]
        ws4[f"H{r4}"].font = font_body
        ws4[f"H{r4}"].fill = fill_white
        ws4[f"H{r4}"].alignment = align_top_left
        ws4[f"H{r4}"].border = border_cell
        r4 += 1

    # -------------------------------------------------------------
    # SHEET 5: 小学生でもわかる解説集 (Analogies & History)
    # -------------------------------------------------------------
    ws5 = wb.create_sheet(title="小学生でもわかる解説集")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:H2")
    ws5["A1"] = "小学生でも直感的にわかる！ API認証の歴史とたとえ話ストーリー"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_title
    ws5["A1"].alignment = align_center

    # Section: Evolution of API Auth
    ws5.merge_cells("A4:H4")
    ws5["A4"] = "第1章：なぜAPI認証はこんなに進化してきたの？（認証の歴史物語）"
    ws5["A4"].font = font_sec_title
    ws5["A4"].fill = fill_sec
    ws5["A4"].alignment = align_left

    history_stories = [
        ("第1世代：パスワードをそのまま送る時代（Basic認証）",
         "【たとえ話】：手紙に「僕の名前はたろう、家の鍵の暗証番号は1234です」と毎回書いて送るようなもの。\n"
         "【どうして困ったの？】：手紙を盗み見されたら、家の鍵を丸ごと乗っ取られてしまいます。相手を信用して鍵を渡すしかなく、とても危険でした。"),
        ("第2世代：秘密の合言葉を決める時代（API-Key認証）",
         "【たとえ話】：本物の鍵の代わりに「ひらけごま」という合言葉（APIキー）を紙に書いて、社内の通用口の警備員に見せるやり方。\n"
         "【いいところ】：簡単で素早い！社内の人しか通らない廊下（イントラネット）ならこれで十分です。\n"
         "【困ったこと】：合言葉を書いた紙を廊下に落としたら、拾った人が誰でも入れてしまいます。合言葉を変えると、全員の合言葉を一斉に変えなきゃいけなくて大変でした。"),
        ("第3世代：時間制限つきの魔法のチケット（OAuth 2.0 と JWT）",
         "【たとえ話】：テーマパークの「1時間乗り放題の腕バンド（アクセストークン）」を発行してもらうやり方。\n"
         "【仕組み】：受付窓口（認可サーバー）で身分を証明すると、「15時30分まで有効」と書かれたスタンプ付きの腕バンドをくれます。アトラクションの係員（API）は、そのスタンプを見るだけで乗せてくれます。\n"
         "【すごいところ】：万が一腕バンドを落としても、1時間経てばただの紙切れになります。本物のパスワードをアトラクションの係員に教える必要もありません！")
    ]

    r5 = 5
    for title, story in history_stories:
        ws5.merge_cells(f"A{r5}:B{r5}")
        ws5[f"A{r5}"] = title
        ws5[f"A{r5}"].font = font_body_bold
        ws5[f"A{r5}"].fill = fill_th_light
        ws5[f"A{r5}"].alignment = align_top_left
        ws5[f"A{r5}"].border = border_cell

        ws5.merge_cells(f"C{r5}:H{r5}")
        ws5[f"C{r5}"] = story
        ws5[f"C{r5}"].font = font_body
        ws5[f"C{r5}"].fill = fill_white
        ws5[f"C{r5}"].alignment = align_top_left
        ws5[f"C{r5}"].border = border_cell
        r5 += 1

    # Section: Character Analogies for our Questions
    r5 += 2
    ws5.merge_cells(f"A{r5}:H{r5}")
    ws5[f"A{r5}"] = "第2章：今回の疑問をすべて「テーマパーク」と「お城」でたとえてみよう！"
    ws5[f"A{r5}"].font = font_sec_title
    ws5[f"A{r5}"].fill = fill_sec
    ws5[f"A{r5}"].alignment = align_left
    r5 += 1

    analogies = [
        ("疑問1：イントラネットAPIはなぜAPI-Keyでいいの？",
         "【たとえ】：社員専用の休憩室のドアノブです。\n"
         "会社の中（イントラネット）には怪しい部外部員は入れない前提なので、社員みんなが知っている暗証番号（API-Key）をポチッと押すだけでドアが開くのが一番ラクで早いからです。"),
        
        ("疑問2：Salesforceのチケットは誰が配っているの？",
         "【たとえ】：Salesforce城の宝物庫に入るための通行手形です。\n"
         "Salesforce城の宝物はSalesforceの王様が管理しているので、通行手形を発行するのは絶対に『Salesforce城の門番（Salesforceのログイン窓口）』です。他の人が勝手に作った手形では門を通れません。"),
        
        ("疑問3：インターネットAPIのチケット売り場はどこにあるの？",
         "【たとえ】：遊園地のアトラクションの看板には『チケット売り場はこちら』と書いてないだけです。\n"
         "ジェットコースターの看板には『身長120cm以上』としか書いてありません。でも、チケットなしで乗れるわけがありませんよね？ 入口の外にあるチケット売り場（Cognito）でちゃんと腕バンド（アクセストークン）を買ってから並ぶのがルールです。"),
        
        ("疑問4：チケットは1回乗るたびに捨てなきゃダメ？（再利用できる？）",
         "【たとえ】：『1時間乗り放題パス』なので、1時間の中なら何回でもアトラクションに乗れます！\n"
         "乗るたびにわざわざ受付に戻って並び直していたら、受付もパンクするし自分も全然遊べませんよね。だから腕バンドは時間いっぱい使い回すのが一番賢いのです。"),
        
        ("疑問5：ロボット同士の通信（M2M）で、IDカードが要らないのはなぜ？",
         "【たとえ】：ロボットには顔も名前も誕生日もないからです。\n"
         "人間なら『私の名前はたろうです』という名札（IDトークン）が必要ですが、工場でロボットがベルトコンベアを動かすのに名札はいりません。『動かす権限がある通行証（アクセストークン）』だけをピッとかざせばお仕事ができるのです。")
    ]

    for q_title, a_story in analogies:
        ws5.merge_cells(f"A{r5}:B{r5}")
        ws5[f"A{r5}"] = q_title
        ws5[f"A{r5}"].font = font_body_bold
        ws5[f"A{r5}"].fill = fill_th_light
        ws5[f"A{r5}"].alignment = align_top_left
        ws5[f"A{r5}"].border = border_cell

        ws5.merge_cells(f"C{r5}:H{r5}")
        ws5[f"C{r5}"] = a_story
        ws5[f"C{r5}"].font = font_body
        ws5[f"C{r5}"].fill = fill_white
        ws5[f"C{r5}"].alignment = align_top_left
        ws5[f"C{r5}"].border = border_cell
        r5 += 1

    # Ensure Meiryo UI is explicitly applied to every cell in the used range
    for ws in wb.worksheets:
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                if cell.font:
                    if cell.font.name != font_name and cell.font.name != "Consolas":
                        cell.font = Font(name=font_name, size=cell.font.size if cell.font.size else 9.5, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color)
                else:
                    cell.font = Font(name=font_name, size=9.5)
            
        # Specific column widths per sheet
        if ws.title == "表紙・全体概要":
            widths = {"A": 4, "B": 6, "C": 24, "D": 26, "E": 14, "F": 14, "G": 14, "H": 14, "I": 14, "J": 14}
        elif ws.title == "API認証方式の比較":
            widths = {"A": 16, "B": 16, "C": 20, "D": 20, "E": 22, "F": 22, "G": 22, "H": 22}
        elif ws.title == "トークンの再利用とTTL":
            widths = {"A": 18, "B": 18, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20}
        elif ws.title == "Cognito_M2M設定解説":
            widths = {"A": 16, "B": 16, "C": 15, "D": 15, "E": 15, "F": 24, "G": 24, "H": 24, "I": 24}
        elif ws.title == "小学生でもわかる解説集":
            widths = {"A": 18, "B": 18, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20, "H": 20}
        else:
            widths = {}

        for c_let, w in widths.items():
            ws.column_dimensions[c_let].width = w

    # Remove default sheet if exists
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    output_path = r"C:\Users\taka_\Agy\API_Auth_Agy\API_Authentication_Architecture_Guide.xlsx"
    wb.save(output_path)
    print(f"Excel successfully created at: {output_path}")

if __name__ == "__main__":
    build_workbook()
