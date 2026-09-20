// lib/excel-handler.ts
import ExcelJS from 'exceljs';

export interface SpecDocsData {
  screen: {
    id: string;
    name: string;
    fields: Array<{
      no: number;
      name: string;
      label: string;
      type: string;
      required: boolean;
      note?: string;
      isNew?: boolean;
    }>;
  };
  api: {
    endpoint: string;
    method: string;
    request_body: Array<{
      no: number;
      name: string;
      type: string;
      required: boolean;
      description?: string;
      isNew?: boolean;
    }>;
  };
  db: {
    table: string;
    tableName: string;
    columns: Array<{
      no: number;
      name: string;
      logicalName: string;
      type: string;
      pk: boolean;
      notNull: boolean;
      description?: string;
      isNew?: boolean;
    }>;
  };
  changeLogs: Array<{
    id: string;
    date: string;
    author: string;
    category: string;
    targetSheets: string;
    summary: string;
    status: string;
    isNew?: boolean;
  }>;
}

export interface MultiFileBundle {
  screen: Uint8Array;
  api: Uint8Array;
  db: Uint8Array;
  changeLog: Uint8Array;
}

// 共通ヘッダースタイル
const headerFill: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF1E293B' }, // Slate-800
};

const headerFont: Partial<ExcelJS.Font> = {
  name: 'Arial',
  color: { argb: 'FFFFFFFF' },
  bold: true,
  size: 10,
};

const titleFill: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF334155' }, // Slate-700
};

const borderStyle: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
};

// 新規追加行のハイライト（薄緑）
const highlightedFill: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFDCFCE7' }, // emerald-100
};

const highlightedFont: Partial<ExcelJS.Font> = {
  name: 'Arial',
  color: { argb: 'FF14532D' }, // emerald-900
  bold: true,
  size: 10,
};

// -------------------------------------------------------------
// 単一独立Excelファイル生成ヘルパー群
// -------------------------------------------------------------

/** 独立 画面設計書 (.xlsx) 生成 */
export async function generateSingleScreenWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('画面設計書_SCR001');
  ws.views = [{ showGridLines: true }];
  ws.mergeCells('A1:F1');
  const titleCell = ws.getCell('A1');
  titleCell.value = '【画面設計書】 画面ID: SCR_001 | 画面名: ユーザー会員登録画面';
  titleCell.fill = titleFill;
  titleCell.font = { ...headerFont, size: 11 };
  titleCell.alignment = { vertical: 'middle', indent: 1 };
  ws.getRow(1).height = 26;

  ws.getRow(2).values = ['No', '項目物理名', '項目論理名', '入力タイプ', '必須', '備考'];
  ws.getRow(2).height = 22;
  ws.getRow(2).eachCell((c) => {
    c.fill = headerFill;
    c.font = headerFont;
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  ws.columns = [
    { key: 'no', width: 6 },
    { key: 'name', width: 22 },
    { key: 'label', width: 22 },
    { key: 'type', width: 16 },
    { key: 'required', width: 10 },
    { key: 'note', width: 35 },
  ];

  ws.addRow([1, 'email', 'メールアドレス', 'email', '○', 'RFC準拠メール形式バリデーション']);
  ws.addRow([2, 'password', 'パスワード', 'password', '○', '8文字以上、英大文字・小文字・記号']);

  ws.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.eachCell((cell) => { cell.border = borderStyle; });
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/** 独立 API仕様書 (.xlsx) 生成 */
export async function generateSingleApiWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('API仕様書_users');
  ws.views = [{ showGridLines: true }];
  ws.mergeCells('A1:E1');
  const titleCell = ws.getCell('A1');
  titleCell.value = '【API仕様書】 エンドポイント: /api/v1/users | HTTPメソッド: POST';
  titleCell.fill = titleFill;
  titleCell.font = { ...headerFont, size: 11 };
  titleCell.alignment = { vertical: 'middle', indent: 1 };
  ws.getRow(1).height = 26;

  ws.getRow(2).values = ['No', 'パラメータ物理名', 'データ型', '必須', '説明'];
  ws.getRow(2).height = 22;
  ws.getRow(2).eachCell((c) => {
    c.fill = headerFill;
    c.font = headerFont;
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  ws.columns = [
    { key: 'no', width: 6 },
    { key: 'name', width: 24 },
    { key: 'type', width: 16 },
    { key: 'required', width: 10 },
    { key: 'description', width: 40 },
  ];

  ws.addRow([1, 'email', 'string', '○', '一意制約あり。重複時409エラー']);
  ws.addRow([2, 'password', 'string', '○', '平文禁止。Argon2idにてハッシュ化']);

  ws.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.eachCell((cell) => { cell.border = borderStyle; });
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/** 独立 テーブル定義書 (.xlsx) 生成 */
export async function generateSingleDbWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('DB定義書_users');
  ws.views = [{ showGridLines: true }];
  ws.mergeCells('A1:G1');
  const titleCell = ws.getCell('A1');
  titleCell.value = '【テーブル定義書】 テーブル名: users | 論理名: ユーザーマスタ';
  titleCell.fill = titleFill;
  titleCell.font = { ...headerFont, size: 11 };
  titleCell.alignment = { vertical: 'middle', indent: 1 };
  ws.getRow(1).height = 26;

  ws.getRow(2).values = ['No', 'カラム物理名', 'カラム論理名', 'データ型', 'PK', 'NOT NULL', '説明'];
  ws.getRow(2).height = 22;
  ws.getRow(2).eachCell((c) => {
    c.fill = headerFill;
    c.font = headerFont;
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  ws.columns = [
    { key: 'no', width: 6 },
    { key: 'name', width: 22 },
    { key: 'logicalName', width: 22 },
    { key: 'type', width: 18 },
    { key: 'pk', width: 8 },
    { key: 'notNull', width: 12 },
    { key: 'description', width: 35 },
  ];

  ws.addRow([1, 'id', 'ユーザーID', 'UUID', '○', '○', '主キー (UUIDv7自動生成)']);
  ws.addRow([2, 'email', 'メールアドレス', 'VARCHAR(255)', '-', '○', '一意インデックス (UNIQUE)']);
  ws.addRow([3, 'password_hash', 'パスワードハッシュ', 'VARCHAR(255)', '-', '○', 'ハッシュ文字列']);
  ws.addRow([4, 'created_at', '作成日時', 'TIMESTAMP', '-', '○', 'DEFAULT CURRENT_TIMESTAMP']);

  ws.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.eachCell((cell) => { cell.border = borderStyle; });
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/** 独立 変更管理簿 (.xlsx) 生成 */
export async function generateSingleChangeLogWorkbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('変更管理簿');
  ws.views = [{ showGridLines: true }];
  ws.columns = [
    { header: 'No', key: 'no', width: 8 },
    { header: '改版日', key: 'date', width: 14 },
    { header: '変更者', key: 'author', width: 18 },
    { header: '区分', key: 'category', width: 12 },
    { header: '対象シート/ファイル', key: 'targetSheets', width: 25 },
    { header: '変更理由・概要', key: 'summary', width: 45 },
    { header: '承認ステータス', key: 'status', width: 15 },
  ];

  const logHeaderRow = ws.getRow(1);
  logHeaderRow.height = 24;
  logHeaderRow.eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  ws.addRow({
    id: 1,
    date: '2026-09-15',
    author: '初期設計者',
    category: '新規作成',
    targetSheets: '全ファイル',
    summary: '会員登録機能 初期仕様策定 (ベースライン確定)',
    status: '承認済',
  });

  ws.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      row.eachCell((cell) => { cell.border = borderStyle; });
    }
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

/** 4つの独立した設計書Excelファイルを一式生成 */
export async function generateMultiFileSpecBundle(): Promise<MultiFileBundle> {
  const [screen, api, db, changeLog] = await Promise.all([
    generateSingleScreenWorkbook(),
    generateSingleApiWorkbook(),
    generateSingleDbWorkbook(),
    generateSingleChangeLogWorkbook(),
  ]);
  return { screen, api, db, changeLog };
}

// -------------------------------------------------------------
// 単一統合ブック（4シート構成）生成
// -------------------------------------------------------------

/**
 * 初期サンプル設計書Excel（4シート構成）を生成してBufferで返す
 */
export async function generateSampleExcelWorkbook(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'DiffSync AI Agent';
  workbook.created = new Date();

  // 1. 【変更管理簿】シート
  const logSheet = workbook.addWorksheet('変更管理簿');
  logSheet.views = [{ showGridLines: true }];
  logSheet.columns = [
    { header: 'No', key: 'no', width: 8 },
    { header: '改版日', key: 'date', width: 14 },
    { header: '変更者', key: 'author', width: 18 },
    { header: '区分', key: 'category', width: 12 },
    { header: '対象シート', key: 'targetSheets', width: 25 },
    { header: '変更理由・概要', key: 'summary', width: 45 },
    { header: '承認ステータス', key: 'status', width: 15 },
  ];

  const logHeaderRow = logSheet.getRow(1);
  logHeaderRow.height = 24;
  logHeaderRow.eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  logSheet.addRow({
    id: 1,
    date: '2026-09-15',
    author: '初期設計者',
    category: '新規作成',
    targetSheets: '全シート',
    summary: '会員登録機能 初期仕様策定 (ベースライン確定)',
    status: '承認済',
  });

  // 2. 【画面設計書】SCR_001
  const screenSheet = workbook.addWorksheet('画面設計書_SCR001');
  screenSheet.views = [{ showGridLines: true }];
  screenSheet.mergeCells('A1:F1');
  const screenTitleCell = screenSheet.getCell('A1');
  screenTitleCell.value = '【画面設計書】 画面ID: SCR_001 | 画面名: ユーザー会員登録画面';
  screenTitleCell.fill = titleFill;
  screenTitleCell.font = { ...headerFont, size: 11 };
  screenTitleCell.alignment = { vertical: 'middle', indent: 1 };
  screenSheet.getRow(1).height = 26;

  screenSheet.getRow(2).values = ['No', '項目物理名', '項目論理名', '入力タイプ', '必須', '備考'];
  screenSheet.getRow(2).height = 22;
  screenSheet.getRow(2).eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  screenSheet.columns = [
    { key: 'no', width: 6 },
    { key: 'name', width: 22 },
    { key: 'label', width: 22 },
    { key: 'type', width: 16 },
    { key: 'required', width: 10 },
    { key: 'note', width: 35 },
  ];

  screenSheet.addRow([1, 'email', 'メールアドレス', 'email', '○', 'RFC準拠メール形式バリデーション']);
  screenSheet.addRow([2, 'password', 'パスワード', 'password', '○', '8文字以上、英大文字・小文字・記号']);

  // 3. 【API仕様書】API_001
  const apiSheet = workbook.addWorksheet('API仕様書_users');
  apiSheet.views = [{ showGridLines: true }];
  apiSheet.mergeCells('A1:E1');
  const apiTitleCell = apiSheet.getCell('A1');
  apiTitleCell.value = '【API仕様書】 エンドポイント: /api/v1/users | HTTPメソッド: POST';
  apiTitleCell.fill = titleFill;
  apiTitleCell.font = { ...headerFont, size: 11 };
  apiTitleCell.alignment = { vertical: 'middle', indent: 1 };
  apiSheet.getRow(1).height = 26;

  apiSheet.getRow(2).values = ['No', 'パラメータ物理名', 'データ型', '必須', '説明'];
  apiSheet.getRow(2).height = 22;
  apiSheet.getRow(2).eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  apiSheet.columns = [
    { key: 'no', width: 6 },
    { key: 'name', width: 24 },
    { key: 'type', width: 16 },
    { key: 'required', width: 10 },
    { key: 'description', width: 40 },
  ];

  apiSheet.addRow([1, 'email', 'string', '○', '一意制約あり。重複時409エラー']);
  apiSheet.addRow([2, 'password', 'string', '○', '平文禁止。Argon2idにてハッシュ化']);

  // 4. 【DB定義書】users
  const dbSheet = workbook.addWorksheet('DB定義書_users');
  dbSheet.views = [{ showGridLines: true }];
  dbSheet.mergeCells('A1:G1');
  const dbTitleCell = dbSheet.getCell('A1');
  dbTitleCell.value = '【テーブル定義書】 テーブル名: users | 論理名: ユーザーマスタ';
  dbTitleCell.fill = titleFill;
  dbTitleCell.font = { ...headerFont, size: 11 };
  dbTitleCell.alignment = { vertical: 'middle', indent: 1 };
  dbSheet.getRow(1).height = 26;

  dbSheet.getRow(2).values = ['No', 'カラム物理名', 'カラム論理名', 'データ型', 'PK', 'NOT NULL', '説明'];
  dbSheet.getRow(2).height = 22;
  dbSheet.getRow(2).eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  dbSheet.columns = [
    { key: 'no', width: 6 },
    { key: 'name', width: 22 },
    { key: 'logicalName', width: 22 },
    { key: 'type', width: 18 },
    { key: 'pk', width: 8 },
    { key: 'notNull', width: 12 },
    { key: 'description', width: 35 },
  ];

  dbSheet.addRow([1, 'id', 'ユーザーID', 'UUID', '○', '○', '主キー (UUIDv7自動生成)']);
  dbSheet.addRow([2, 'email', 'メールアドレス', 'VARCHAR(255)', '-', '○', '一意インデックス (UNIQUE)']);
  dbSheet.addRow([3, 'password_hash', 'パスワードハッシュ', 'VARCHAR(255)', '-', '○', 'ハッシュ文字列']);
  dbSheet.addRow([4, 'created_at', '作成日時', 'TIMESTAMP', '-', '○', 'DEFAULT CURRENT_TIMESTAMP']);

  // 罫線設定
  [logSheet, screenSheet, apiSheet, dbSheet].forEach((sheet) => {
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        row.eachCell((cell) => {
          cell.border = borderStyle;
        });
      }
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer as ArrayBuffer);
}

// -------------------------------------------------------------
// Excelバイナリ パース処理
// -------------------------------------------------------------

/**
 * 統合または個別Excelバイナリから設計データを構造化JSONとしてパースする
 */
export async function parseExcelWorkbook(buffer: ArrayBuffer | Uint8Array): Promise<SpecDocsData> {
  const workbook = new ExcelJS.Workbook();
  const nodeBuffer = Buffer.from(buffer as any);
  await workbook.xlsx.load(nodeBuffer as any);

  let screenSheet = workbook.worksheets.find((s) => s.name.includes('画面') || s.name.toLowerCase().includes('screen'));
  let apiSheet = workbook.worksheets.find((s) => s.name.includes('API') || s.name.toLowerCase().includes('api'));
  let dbSheet = workbook.worksheets.find((s) => s.name.includes('DB') || s.name.includes('テーブル') || s.name.toLowerCase().includes('db'));
  let logSheet = workbook.worksheets.find((s) => s.name.includes('変更管理') || s.name.toLowerCase().includes('change'));

  // フォールバック（順序で判定）
  if (!logSheet && workbook.worksheets.length > 0) logSheet = workbook.worksheets[0];
  if (!screenSheet && workbook.worksheets.length > 1) screenSheet = workbook.worksheets[1];
  if (!apiSheet && workbook.worksheets.length > 2) apiSheet = workbook.worksheets[2];
  if (!dbSheet && workbook.worksheets.length > 3) dbSheet = workbook.worksheets[3];

  // 1. 変更管理簿の抽出
  const changeLogs: SpecDocsData['changeLogs'] = [];
  if (logSheet) {
    let emptyCount = 0;
    for (let r = 2; r <= 500; r++) {
      const row = logSheet.getRow(r);
      const rowValues = row.values as any[];
      if (rowValues && rowValues[1] && String(rowValues[1]).trim() !== '') {
        emptyCount = 0;
        changeLogs.push({
          id: `CHG-${String(rowValues[1]).padStart(3, '0')}`,
          date: String(rowValues[2] || ''),
          author: String(rowValues[3] || ''),
          category: String(rowValues[4] || ''),
          targetSheets: String(rowValues[5] || ''),
          summary: String(rowValues[6] || ''),
          status: String(rowValues[7] || '承認済'),
        });
      } else {
        emptyCount++;
        if (emptyCount >= 5) break;
      }
    }
  }

  // 2. 画面設計書の抽出
  const screenFields: SpecDocsData['screen']['fields'] = [];
  let screenId = 'SCR_001';
  let screenName = 'ユーザー会員登録画面';
  if (screenSheet) {
    const titleVal = String(screenSheet.getCell('A1').value || '');
    if (titleVal.includes('SCR_')) {
      const match = titleVal.match(/SCR_[A-Za-z0-9_]+/);
      if (match) screenId = match[0];
    }
    let emptyCount = 0;
    for (let r = 3; r <= 500; r++) {
      const row = screenSheet.getRow(r);
      const v = row.values as any[];
      if (v && v[2] && String(v[2]).trim() !== '') {
        emptyCount = 0;
        screenFields.push({
          no: Number(v[1]) || r - 2,
          name: String(v[2]),
          label: String(v[3] || v[2]),
          type: String(v[4] || 'text'),
          required: String(v[5]).includes('○') || String(v[5]).toLowerCase() === 'true',
          note: String(v[6] || ''),
        });
      } else {
        emptyCount++;
        if (emptyCount >= 5) break;
      }
    }
  }

  // 3. API仕様書の抽出
  const apiParams: SpecDocsData['api']['request_body'] = [];
  let endpoint = '/api/v1/users';
  let method = 'POST';
  if (apiSheet) {
    const titleVal = String(apiSheet.getCell('A1').value || '');
    if (titleVal.includes('/api/')) {
      const match = titleVal.match(/\/api\/[^\s|]+/);
      if (match) endpoint = match[0];
    }
    let emptyCount = 0;
    for (let r = 3; r <= 500; r++) {
      const row = apiSheet.getRow(r);
      const v = row.values as any[];
      if (v && v[2] && String(v[2]).trim() !== '') {
        emptyCount = 0;
        apiParams.push({
          no: Number(v[1]) || r - 2,
          name: String(v[2]),
          type: String(v[3] || 'string'),
          required: String(v[4]).includes('○') || String(v[4]).toLowerCase() === 'true',
          description: String(v[5] || ''),
        });
      } else {
        emptyCount++;
        if (emptyCount >= 5) break;
      }
    }
  }

  // 4. DB定義書の抽出
  const dbColumns: SpecDocsData['db']['columns'] = [];
  let tableName = 'users';
  let tableLogicalName = 'ユーザーマスタ';
  if (dbSheet) {
    const titleVal = String(dbSheet.getCell('A1').value || '');
    if (titleVal.includes('テーブル名:')) {
      const match = titleVal.match(/テーブル名:\s*([^\s|]+)/);
      if (match) tableName = match[1];
    }
    let emptyCount = 0;
    for (let r = 3; r <= 500; r++) {
      const row = dbSheet.getRow(r);
      const v = row.values as any[];
      if (v && v[2] && String(v[2]).trim() !== '') {
        emptyCount = 0;
        dbColumns.push({
          no: Number(v[1]) || r - 2,
          name: String(v[2]),
          logicalName: String(v[3] || v[2]),
          type: String(v[4] || 'VARCHAR(255)'),
          pk: String(v[5]).includes('○'),
          notNull: String(v[6]).includes('○'),
          description: String(v[7] || ''),
        });
      } else {
        emptyCount++;
        if (emptyCount >= 5) break;
      }
    }
  }

  return {
    screen: { id: screenId, name: screenName, fields: screenFields },
    api: { endpoint, method, request_body: apiParams },
    db: { table: tableName, tableName: tableLogicalName, columns: dbColumns },
    changeLogs: changeLogs.length > 0 ? changeLogs : [
      {
        id: 'CHG-001',
        date: '2026-09-15',
        author: '初期設計者',
        category: '新規作成',
        targetSheets: '全シート',
        summary: '初期ベースライン策定',
        status: '承認済',
      },
    ],
  };
}

/**
 * 4つの独立した各Excelバイナリをまとめてパースして統合データを返す
 */
export async function parseMultiFileSpecs(files: {
  screen?: ArrayBuffer | Uint8Array;
  api?: ArrayBuffer | Uint8Array;
  db?: ArrayBuffer | Uint8Array;
  changeLog?: ArrayBuffer | Uint8Array;
}): Promise<SpecDocsData> {
  const [screenData, apiData, dbData, logData] = await Promise.all([
    files.screen ? parseExcelWorkbook(files.screen) : null,
    files.api ? parseExcelWorkbook(files.api) : null,
    files.db ? parseExcelWorkbook(files.db) : null,
    files.changeLog ? parseExcelWorkbook(files.changeLog) : null,
  ]);

  return {
    screen: screenData?.screen || { id: 'SCR_001', name: 'ユーザー会員登録画面', fields: [] },
    api: apiData?.api || { endpoint: '/api/v1/users', method: 'POST', request_body: [] },
    db: dbData?.db || { table: 'users', tableName: 'ユーザーマスタ', columns: [] },
    changeLogs: logData?.changeLogs || screenData?.changeLogs || [],
  };
}

// -------------------------------------------------------------
// Excelバイナリ パッチ適用処理
// -------------------------------------------------------------

/**
 * 単一統合ワークブック（4シート構成）にパッチを適用
 */

  // 削除処理 (下から上へ走査して行を削除)
  const applyDeletes = (sheet: ExcelJS.Worksheet | undefined, toDeleteList: string[] | undefined, nameColIndex: number, startRowIndex: number) => {
    if (!sheet || !toDeleteList || toDeleteList.length === 0) return;
    const rowsToDelete: number[] = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > startRowIndex) {
        const v = row.values as any[];
        if (v && v[nameColIndex] && toDeleteList.includes(String(v[nameColIndex]))) {
          rowsToDelete.push(rowNumber);
        }
      }
    });
    // 下から上に削除しないとインデックスがずれる
    rowsToDelete.reverse().forEach((rowNum) => {
      sheet.spliceRows(rowNum, 1);
    });
  };


function findFirstEmptyRowIndex(sheet: ExcelJS.Worksheet, nameColIndex: number, startRow: number): number {
  for (let r = startRow; r <= 500; r++) {
    const row = sheet.getRow(r);
    const v = row.values as any[];
    if (!v || !v[nameColIndex] || String(v[nameColIndex]).trim() === '') {
      return r;
    }
  }
  return startRow;
}

export async function applyPatchToExcelWorkbook(
  buffer: ArrayBuffer | Uint8Array,
  patch: {
    screenFieldsToAdd?: Array<{ name: string; label: string; type: string; required: boolean; note: string }>;
    screenFieldsToDelete?: string[];
    apiParamsToAdd?: Array<{ name: string; type: string; required: boolean; description: string }>;
    apiParamsToDelete?: string[];
    dbColumnsToAdd?: Array<{ name: string; logicalName: string; type: string; pk?: boolean; notNull?: boolean; description: string }>;
    dbColumnsToDelete?: string[];
  },
  newLogRecord: {
    ticketId: string;
    date: string;
    author: string;
    category: string;
    targetSheets: string;
    summary: string;
    status: string;
  }
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  const nodeBuffer = Buffer.from(buffer as any);
  await workbook.xlsx.load(nodeBuffer as any);

  let screenSheet = workbook.worksheets.find((s) => s.name.includes('画面') || s.name.toLowerCase().includes('screen'));
  let apiSheet = workbook.worksheets.find((s) => s.name.includes('API') || s.name.toLowerCase().includes('api'));
  let dbSheet = workbook.worksheets.find((s) => s.name.includes('DB') || s.name.includes('テーブル') || s.name.toLowerCase().includes('db'));
  let logSheet = workbook.worksheets.find((s) => s.name.includes('変更管理') || s.name.toLowerCase().includes('change'));


  applyDeletes(screenSheet, patch.screenFieldsToDelete, 2, 1);
  applyDeletes(apiSheet, patch.apiParamsToDelete, 2, 2);
  applyDeletes(dbSheet, patch.dbColumnsToDelete, 2, 2);


  applyDeletes(screenSheet, patch.screenFieldsToDelete, 2, 1);
  applyDeletes(apiSheet, patch.apiParamsToDelete, 2, 2);
  applyDeletes(dbSheet, patch.dbColumnsToDelete, 2, 2);

  // 1. 画面設計書シートへの行追加
  if (screenSheet && patch.screenFieldsToAdd && patch.screenFieldsToAdd.length > 0) {
    let currentMaxNo = 0;
    for (let r = 3; r <= 500; r++) {
      const row = screenSheet.getRow(r);
      const v = row.values as any[];
      if (v && v[2] && String(v[2]).trim() !== '') {
        if (Number(v[1]) > currentMaxNo) currentMaxNo = Number(v[1]);
      } else {
        break;
      }
    }

    let insertRowIndex = findFirstEmptyRowIndex(screenSheet, 2, 3);
    for (const f of patch.screenFieldsToAdd) {
      currentMaxNo++;
      const newRow = screenSheet.getRow(insertRowIndex);
      newRow.values = [
        currentMaxNo,
        f.name,
        f.label,
        f.type,
        f.required ? '○' : '-',
        f.note || '',
      ];
      newRow.height = 22;
      newRow.eachCell((cell) => {
        cell.border = borderStyle;
      });
      insertRowIndex++;
    }
  }

  // 2. API仕様書シートへの行追加
  if (apiSheet && patch.apiParamsToAdd && patch.apiParamsToAdd.length > 0) {
    let currentMaxNo = 0;
    for (let r = 3; r <= 500; r++) {
      const row = apiSheet.getRow(r);
      const v = row.values as any[];
      if (v && v[2] && String(v[2]).trim() !== '') {
        if (Number(v[1]) > currentMaxNo) currentMaxNo = Number(v[1]);
      } else {
        break;
      }
    }

    let insertRowIndex = findFirstEmptyRowIndex(apiSheet, 2, 3);
    for (const p of patch.apiParamsToAdd) {
      currentMaxNo++;
      const newRow = apiSheet.getRow(insertRowIndex);
      newRow.values = [
        currentMaxNo,
        p.name,
        p.type,
        p.required ? '○' : '-',
        p.description || '',
      ];
      newRow.height = 22;
      newRow.eachCell((cell) => {
        cell.border = borderStyle;
      });
      insertRowIndex++;
    }
  }

  // 3. DB定義書シートへの行追加
  if (dbSheet && patch.dbColumnsToAdd && patch.dbColumnsToAdd.length > 0) {
    let currentMaxNo = 0;
    for (let r = 3; r <= 500; r++) {
      const row = dbSheet.getRow(r);
      const v = row.values as any[];
      if (v && v[2] && String(v[2]).trim() !== '') {
        if (Number(v[1]) > currentMaxNo) currentMaxNo = Number(v[1]);
      } else {
        break;
      }
    }

    let insertRowIndex = findFirstEmptyRowIndex(dbSheet, 2, 3);
    for (const c of patch.dbColumnsToAdd) {
      currentMaxNo++;
      const newRow = dbSheet.getRow(insertRowIndex);
      newRow.values = [
        currentMaxNo,
        c.name,
        c.logicalName,
        c.type,
        c.pk ? '○' : '-',
        c.notNull ? '○' : '-',
        c.description || '',
      ];
      newRow.height = 22;
      newRow.eachCell((cell) => {
        cell.border = borderStyle;
      });
      insertRowIndex++;
    }
  }

  // 4. 変更管理簿シートへの自律追記
  if (logSheet) {
    let currentMaxNo = 0;
    for (let r = 2; r <= 500; r++) {
      const row = logSheet.getRow(r);
      const v = row.values as any[];
      if (v && v[1] && String(v[1]).trim() !== '') {
        if (Number(v[1]) > currentMaxNo) currentMaxNo = Number(v[1]);
      } else {
        break;
      }
    }

    let insertRowIndex = findFirstEmptyRowIndex(logSheet, 1, 2);
    const newLogRow = logSheet.getRow(insertRowIndex);
    newLogRow.values = [
      currentMaxNo + 1,
      newLogRecord.date,
      newLogRecord.author || 'AIエージェント',
      newLogRecord.category || '仕様変更',
      newLogRecord.targetSheets,
      newLogRecord.summary,
      newLogRecord.status || '承認待機',
    ];
    newLogRow.height = 24;
    newLogRow.eachCell((cell) => {
      cell.border = borderStyle;
    });
  }

  const updatedBuffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(updatedBuffer as ArrayBuffer);
}

/**
 * 4つの独立した各Excelファイルに串刺しで直接パッチを適用し、更新後の4ファイルバイナリを返す
 */
export async function applyPatchToMultiFiles(
  files: MultiFileBundle,
  patch: {
    screenFieldsToAdd?: Array<{ name: string; label: string; type: string; required: boolean; note: string }>;
    screenFieldsToDelete?: string[];
    apiParamsToAdd?: Array<{ name: string; type: string; required: boolean; description: string }>;
    apiParamsToDelete?: string[];
    dbColumnsToAdd?: Array<{ name: string; logicalName: string; type: string; pk?: boolean; notNull?: boolean; description: string }>;
    dbColumnsToDelete?: string[];
  },
  newLogRecord: {
    ticketId: string;
    date: string;
    author: string;
    category: string;
    targetSheets: string;
    summary: string;
    status: string;
  }
): Promise<MultiFileBundle> {
  // 各ファイルを個別に読み込んで直接更新
  const [updatedScreen, updatedApi, updatedDb, updatedLog] = await Promise.all([
    // 1. 画面設計書.xlsx 直接更新
    applyPatchToExcelWorkbook(files.screen, { screenFieldsToAdd: patch.screenFieldsToAdd, screenFieldsToDelete: patch.screenFieldsToDelete }, newLogRecord),
    // 2. API仕様書.xlsx 直接更新
    applyPatchToExcelWorkbook(files.api, { apiParamsToAdd: patch.apiParamsToAdd, apiParamsToDelete: patch.apiParamsToDelete }, newLogRecord),
    // 3. テーブル定義書.xlsx 直接更新
    applyPatchToExcelWorkbook(files.db, { dbColumnsToAdd: patch.dbColumnsToAdd, dbColumnsToDelete: patch.dbColumnsToDelete }, newLogRecord),
    // 4. 変更管理簿.xlsx 直接更新
    applyPatchToExcelWorkbook(files.changeLog, {}, newLogRecord),
  ]);

  return {
    screen: updatedScreen,
    api: updatedApi,
    db: updatedDb,
    changeLog: updatedLog,
  };
}

