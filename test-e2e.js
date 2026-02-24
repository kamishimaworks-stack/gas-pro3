const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ locale: 'ja-JP', viewport: { width: 1400, height: 900 } });
  const page = await context.newPage();

  const consoleErrors = [];
  const jsErrors = [];
  const bugs = [];

  page.on('console', msg => { if (msg.type() === 'error') consoleErrors.push(msg.text()); });
  page.on('pageerror', err => { jsErrors.push(err.message); });

  const filePath = 'file:///' + path.resolve(__dirname, 'index.html').replace(/\\/g, '/');
  console.log('=== ローカルモック E2E テスト ===');

  try {
    // ============ ページ読み込み ============
    console.log('\n[1] ページ読み込み...');
    await page.goto(filePath, { waitUntil: 'load', timeout: 30000 });
    await page.waitForSelector('#app:not([v-cloak])', { timeout: 15000 });
    console.log('  OK: Vue.js マウント完了');
    await page.waitForTimeout(2000);
    await page.screenshot({ path: 'screenshot-01-menu.png', fullPage: true });

    const titles = await page.$$eval('.menu-title', els => els.map(e => e.textContent.trim()));
    console.log('  メニュータイトル:', titles);

    // ============ 御見積書作成画面 ============
    console.log('\n[2] 御見積書作成画面...');
    await page.locator('.menu-title:text("御見積書作成")').first().click();
    await page.waitForTimeout(3000);

    // [2a] 発注先ドロップダウン
    console.log('  [2a] 発注先ドロップダウン...');
    const vendorOpts = await page.locator('td.bg-orange-50 select').first().locator('option').allTextContents();
    console.log('    選択肢:', vendorOpts);
    if (vendorOpts.length <= 1) bugs.push('発注先ドロップダウンに選択肢がない');
    else console.log('    OK:', vendorOpts.length - 1, '件');

    // [2b] 単位ドロップダウン
    console.log('  [2b] 単位ドロップダウン...');
    const unitOpts = await page.locator('select.input-cell.text-center').first().locator('option').allTextContents();
    console.log('    選択肢:', unitOpts);
    if (unitOpts.length <= 1) bugs.push('単位ドロップダウンに選択肢がない');
    else console.log('    OK:', unitOpts.length - 1, '件');

    // [2c] 金額リアルタイム計算
    console.log('  [2c] 金額リアルタイム計算...');
    const qtyInput = page.locator('input[type="number"][data-row="0"][data-col="2"]').first();
    const priceInput = page.locator('input[type="number"][data-row="0"][data-col="4"]').first();
    await qtyInput.click(); await qtyInput.fill('10'); await qtyInput.dispatchEvent('input');
    await page.waitForTimeout(300);
    await priceInput.click(); await priceInput.fill('5000'); await priceInput.dispatchEvent('input');
    await page.waitForTimeout(300);
    const amountText = await page.locator('td.num-cell.font-mono.bg-blue-100').first().textContent();
    console.log('    金額:', amountText.trim());
    if (amountText.includes('50,000')) console.log('    OK: 計算正常');
    else bugs.push('金額リアルタイム計算が不正: ' + amountText.trim());

    // [2d] Enterキーでセル移動テスト（行追加してから）
    console.log('  [2d] Enterキーセル移動...');
    // 行追加ボタンをクリックして2行目を作る
    const addRowBtn = page.locator('button:has-text("+ 行追加")').first();
    if (await addRowBtn.count() > 0) {
      await addRowBtn.click();
      await page.waitForTimeout(500);
    }
    // 1行目(row=0)の数量にフォーカスしてEnter
    await qtyInput.focus();
    await page.waitForTimeout(200);
    const before = await page.evaluate(() => ({ row: document.activeElement?.dataset?.row, col: document.activeElement?.dataset?.col, tag: document.activeElement?.tagName }));
    await page.keyboard.press('Enter');
    await page.waitForTimeout(500);
    const after = await page.evaluate(() => ({ row: document.activeElement?.dataset?.row, col: document.activeElement?.dataset?.col, tag: document.activeElement?.tagName }));
    console.log('    Before:', JSON.stringify(before), '→ After:', JSON.stringify(after));
    if (before.row === after.row && before.col === after.col) {
      bugs.push('Enterキーでセル移動が動作しない (row0→row1に移動すべき)');
    } else {
      console.log('    OK: セル移動成功');
    }

    // [2e] 品名入力＋コピー
    console.log('  [2e] コピーテスト...');
    await page.locator('textarea[placeholder="品名"]').first().fill('テスト足場');
    // ヘッダーも入力
    const clientSelect = page.locator('select').filter({ hasText: 'テスト建設' }).first();
    if (await clientSelect.count() > 0) {
      await clientSelect.selectOption({ label: 'テスト建設' });
    }
    await page.waitForTimeout(300);

    await page.locator('.function-bar .f-btn:has-text("コピー")').click();
    await page.waitForTimeout(500);
    const copyToast = await page.locator('.toast').first().textContent().catch(() => '');
    console.log('    トースト:', copyToast.trim());
    const hasPasteBtn = await page.locator('.function-bar .f-btn:has-text("貼り付け")').count();
    console.log('    貼り付けボタン:', hasPasteBtn > 0 ? 'OK 表示' : 'NG 非表示');
    if (hasPasteBtn === 0) bugs.push('コピー後に貼り付けボタンが表示されない');

    await page.screenshot({ path: 'screenshot-02-edit-filled.png', fullPage: true });

    // ============ メニューに戻る→発注書作成画面 ============
    console.log('\n[3] 発注書作成画面...');
    // 「メニュー」ボタンで戻る
    const menuBtn = page.locator('button:has-text("メニュー")');
    if (await menuBtn.count() > 0) {
      await menuBtn.click();
      await page.waitForTimeout(1000);
      // 未保存確認ダイアログがあれば「保存せず移動」
      const discardBtn = page.locator('button:has-text("保存せず移動")');
      if (await discardBtn.count() > 0) {
        await discardBtn.click();
        await page.waitForTimeout(1000);
      }
    }
    await page.waitForTimeout(1000);
    await page.screenshot({ path: 'screenshot-03-back-to-menu.png', fullPage: true });

    // 発注書作成カードをクリック
    const orderCard = page.locator('.menu-title:text("発注書作成")').first();
    if (await orderCard.count() > 0) {
      await orderCard.click();
      await page.waitForTimeout(3000);
      await page.screenshot({ path: 'screenshot-04-order.png', fullPage: true });
      console.log('  screenshot-04-order.png');

      // [3a] 貼り付け
      console.log('  [3a] 貼り付けテスト...');
      const orderPaste = page.locator('.function-bar .f-btn:has-text("貼り付け")');
      if (await orderPaste.count() > 0) {
        await orderPaste.click();
        await page.waitForTimeout(1000);

        const product = await page.locator('textarea[placeholder="品名"]').first().inputValue().catch(() => '');
        console.log('    品名:', product);
        if (product === 'テスト足場') console.log('    OK: 貼り付け成功');
        else bugs.push('貼り付け後の品名が不正: "' + product + '"');

        const estQty = await page.locator('input[type="number"][data-row="0"][data-col="2"]').first().inputValue().catch(() => '');
        console.log('    見積数量:', estQty);
        if (estQty === '10') console.log('    OK: 数量反映');
        else bugs.push('貼り付け後の数量が不正: "' + estQty + '"');

        await page.screenshot({ path: 'screenshot-05-order-pasted.png', fullPage: true });
      } else {
        bugs.push('発注書画面に貼り付けボタンが表示されない');
      }

      // [3b] 発注書の発注先・単位ドロップダウン
      console.log('  [3b] 発注書ドロップダウン...');
      const ovOpts = await page.locator('td.bg-orange-50 select').first().locator('option').allTextContents().catch(() => []);
      console.log('    発注先:', ovOpts);
      const ouOpts = await page.locator('select.input-cell.text-center').first().locator('option').allTextContents().catch(() => []);
      console.log('    単位:', ouOpts);

    } else {
      bugs.push('メニューに戻った後「発注書作成」カードが見つからない');
    }

    // ============ まとめ ============
    console.log('\n\n========================================');
    console.log('         テスト結果サマリー');
    console.log('========================================');
    const realErrors = jsErrors.filter(e => !e.includes('favicon'));
    const realConsoleErrors = consoleErrors.filter(e => !e.includes('favicon') && !e.includes('ERR_FILE_NOT_FOUND') && !e.includes('net::'));
    console.log('JS例外:', realErrors.length ? realErrors : 'なし');
    console.log('コンソールエラー:', realConsoleErrors.length ? realConsoleErrors : 'なし');
    console.log('');
    if (bugs.length === 0) {
      console.log('バグ: なし - 全テスト通過!');
    } else {
      console.log('検出されたバグ (' + bugs.length + '件):');
      bugs.forEach((b, i) => console.log('  ' + (i+1) + '. ' + b));
    }
    console.log('========================================');

  } catch (e) {
    console.log('FATAL:', e.message);
  }
  await browser.close();
})();
