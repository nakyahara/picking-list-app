// ============================================================
// 既存の PL_writePickingList_FillE_and_ExportPDF() の末尾、
// PDF出力の直後（SpreadsheetApp.flush() の前）に以下を追加
// ============================================================

  /** 6) マッピングCSV出力（商品ID → 納品プランNo） */
  const csvRows = ['商品ID,納品プランNo'];  // ヘッダー
  for (let i = 0; i < bodyAD.length; i++) {
    const productId = bodyAD[i][2];  // C列: 商品ID
    const planNo = colE[i][0];       // E列: 納品プランNo
    if (productId && planNo) {
      // CSVエスケープ
      const safePid = productId.includes(',') ? `"${productId}"` : productId;
      // 改行区切りの複数プランは「 / 」で結合して全て保持
      const allPlans = planNo.split('\n').map(p => p.trim()).filter(p => p).join(' / ');
      const safePlan = allPlans.includes(',') ? `"${allPlans}"` : allPlans;
      csvRows.push(`${safePid},${safePlan}`);
    }
  }
  const csvContent = csvRows.join('\n');
  const csvBlob = Utilities.newBlob(csvContent, 'text/csv',
    `mapping_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`);
  DriveApp.getFolderById(PL_OUTPUT_FOLDER_ID).createFile(csvBlob);

// ============================================================
// 挿入位置の目印:
//
// 既存コード:
//   const pdf = PL_exportSheetToPdf_(...);
//   DriveApp.getFolderById(PL_OUTPUT_FOLDER_ID).createFile(pdf);
//
//   ↓ ここに上記コードを追加 ↓
//
//   (関数の閉じ括弧 } の直前)
// ============================================================
