# マイルストーン2: 商品データのスクレイピングとスプレッドシート保存

## 概要
フィルタリングされた商品リストから、各商品の詳細情報を自動的にスクレイピングし、Googleスプレッドシートに保存します。

## 実装内容

### 1. データ収集項目
以下の情報を各商品から抽出します：
- **ASIN**: Amazon商品識別コード
- **商品名**: 製品の名称
- **商品数**: 数量
- **参考価格**: 割引前の元の価格（JPY）
- **単価**: 現在の販売価格（JPY）
- **割引率**: パーセンテージ（%）
- **割引額**: 金額（JPY）

### 2. スクレイピング手法（スクロール＆スクレイプ方式）
```python
def scrape_all_products(page, context):
    """
    フィルタリングされた結果ページから全商品をスクレイピング
    段階的にスクロールしながら、表示された商品を順次スクレイピング
    各商品を別タブで開き、情報を収集後、タブを閉じる
    """
```

**処理フロー（改良版）:**
1. 現在表示されている商品コンテナを検出（`.sg-product-gallery-item`）
2. 新しい商品のURLを抽出（既にスクレイピング済みのものは除外）
3. 検出された各商品について：
   - 新しいタブで商品詳細ページを開く
   - 必要なデータを抽出（`scrape_product_details()`）
   - タブを閉じる
   - 次の商品へ（0.8秒の待機時間）
4. ページを500pxスクロールダウン
5. 1.5秒待機（遅延読み込み対応）
6. 新しい商品が見つかる限り、1〜5を繰り返す
7. 3回連続で新商品が見つからない場合、または最大20回スクロールで終了

**メリット:**
- メモリ効率的：全商品を一度にロードしない
- より自然な動作：人間のブラウジングパターンに近い
- 遅延読み込み対応：スクロールで動的に読み込まれる商品も取得可能
- 中断可能：エラー発生時も既にスクレイピングした商品データは保持

### 3. データ抽出ロジック

#### ASIN抽出
```python
# URLから抽出: /dp/ASIN/ または /gp/product/ASIN
asin_match = re.search(r'/(?:dp|gp/product)/([A-Z0-9]{10})', product_url)
```

#### 商品名
```python
# セレクタ: #productTitle, .a-truncate-cut
name_elem = page.locator("#productTitle, .a-truncate-cut").first
product_data['name'] = name_elem.inner_text().strip()
```

#### 参考価格（割引前価格）
```python
# セレクタ: .a-price.a-text-price, .basisPrice, [data-a-strike='true']
ref_price_selectors = [
    ".a-price.a-text-price",
    ".basisPrice .a-price .a-offscreen",
    "[data-a-strike='true'] .a-offscreen",
    ".a-text-strike .a-offscreen"
]
```

#### 現在価格（単価）
```python
# セレクタ: .a-price.priceToPay, #priceblock_ourprice
price_selectors = [
    ".a-price.priceToPay .a-offscreen",
    ".a-price .a-offscreen",
    "#priceblock_ourprice",
    "#priceblock_dealprice"
]
```

#### 割引率・割引額の計算
```python
if product_data['reference_price'] and product_data['unit_price']:
    ref = float(product_data['reference_price'])
    curr = float(product_data['unit_price'])
    discount_amount = ref - curr
    discount_rate = (discount_amount / ref) * 100
```

### 4. Googleスプレッドシート統合

#### 認証
```python
def get_sheets_service():
    """
    gspreadを使用してGoogle Sheets APIに認証
    OAuth 2.0フローでトークンを取得・保存
    """
```

**使用スコープ:**
- `https://www.googleapis.com/auth/gmail.readonly`（OTP取得用）
- `https://www.googleapis.com/auth/spreadsheets`（シート書き込み用）

#### データ送信
```python
def send_to_google_sheets(products_data):
    """
    スクレイピングした商品データをGoogleスプレッドシートに送信
    """
```

**処理:**
1. Google Sheets APIで認証
2. スプレッドシートを開く（ID: `1t_HjbOjlcgwZACo2glY8w-OfUVAGa3TEcX2h5wIwejk`）
3. 既存データをクリア
4. ヘッダー行を準備
5. 全商品データを一度に書き込み（`worksheet.update('A1', rows)`）

**スプレッドシート形式:**
```
| ASIN | Product Name | Number of Products | Reference Price (JPY) | Price per Unit (JPY) | Discount Rate (%) | Discount Amount (JPY) |
|------|--------------|--------------------|-----------------------|----------------------|-------------------|-----------------------|
| ...  | ...          | ...                | ...                   | ...                  | ...               | ...                   |
```

### 5. エラーハンドリング
- 各商品のスクレイピング失敗時も処理を継続
- データ抽出できない項目は空文字列で保存
- スプレッドシート送信失敗時はエラーメッセージを表示

### 6. 使用技術
- **Playwright**: ブラウザ自動化、複数タブ管理
- **gspread**: Googleスプレッドシート操作
- **google-auth**: OAuth 2.0認証
- **正規表現**: ASIN抽出、価格の数値変換

### 7. パフォーマンス最適化
- 各商品間に0.8秒の待機時間（サーバー負荷軽減と高速化のバランス）
- スクロール後1.5秒待機（遅延読み込みコンテンツの表示待ち）
- 新しいタブで商品を開き、閉じることでメモリ管理
- 一括書き込み（`worksheet.update()`）でAPI呼び出し回数を削減
- 重複スクレイピング防止：既に処理した商品URLを記録
- 自動終了検知：新商品が見つからなくなったら自動停止

## 実行結果
スクレイピングが完了すると、以下の情報が表示されます：
```
[Scroll X] Found Y new products to scrape
[INFO] Total products discovered so far: Z
  [N] Scraping product...
    [SUCCESS] Scraped: 商品名...
...
[SUCCESS] Scraped X products successfully
[INFO] Total scrolls: Y
[SUCCESS] Successfully wrote X products to Google Sheets!
[INFO] View at: https://docs.google.com/spreadsheets/d/...
[INFO] Closing browser...
[SUCCESS] Browser closed
```

## 注意事項
1. **Google認証**: 初回実行時にブラウザでGoogle認証が必要
2. **スプレッドシートアクセス権**: 使用するGoogleアカウントにスプレッドシートの編集権限が必要
3. **レート制限**: 大量の商品をスクレイピングする場合、待機時間の調整が必要な場合あり
4. **セレクタ変更**: Amazonのページ構造が変更された場合、セレクタの更新が必要
