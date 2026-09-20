export const dynamic = 'force-dynamic';
// app/api/excel/bundle/route.ts
import { NextResponse } from 'next/server';
import { generateMultiFileSpecBundle } from '@/lib/excel-handler';
import fs from 'fs';
import path from 'path';

const DRIVE_DIR = process.env.GOOGLE_DRIVE_MOCK_DIR || path.join(process.cwd(), 'CloudDriveSync');

const FILE_MAP: Record<string, { fileName: string; category: string }> = {
  screen:    { fileName: '画面設計書_SCR001.xlsx', category: 'フロントエンド設計 (UI/UX)' },
  api:       { fileName: 'API仕様書_users.xlsx',  category: 'バックエンド仕様 (API)' },
  db:        { fileName: 'テーブル定義書_users.xlsx', category: 'データベース定義 (DBA)' },
  changeLog: { fileName: '変更管理簿_2026.xlsx',  category: 'プロジェクト管理 (PMO)' },
};

function readFromDisk(key: string): Buffer | null {
  try {
    const filePath = path.join(DRIVE_DIR, FILE_MAP[key].fileName);
    if (fs.existsSync(filePath)) {
      return fs.readFileSync(filePath);
    }
  } catch (_) {}
  return null;
}

export async function GET() {
  try {
    // CloudDriveSync/ に保存済みファイルがあればそちらを優先して読み込む
    const savedScreen    = readFromDisk('screen');
    const savedApi       = readFromDisk('api');
    const savedDb        = readFromDisk('db');
    const savedChangeLog = readFromDisk('changeLog');

    const allSaved = savedScreen && savedApi && savedDb && savedChangeLog;

    // 1件でも欠けている場合はサンプルを生成（部分的に混在させない）
    const bundle = allSaved ? null : await generateMultiFileSpecBundle();

    const toBase64 = (saved: Buffer | null, generated: Uint8Array): string =>
      Buffer.from(saved ?? generated).toString('base64');

    return NextResponse.json({
      success: true,
      source: allSaved ? 'saved' : 'generated',
      files: {
        screen: {
          ...FILE_MAP.screen,
          base64: allSaved
            ? savedScreen!.toString('base64')
            : toBase64(null, bundle!.screen),
        },
        api: {
          ...FILE_MAP.api,
          base64: allSaved
            ? savedApi!.toString('base64')
            : toBase64(null, bundle!.api),
        },
        db: {
          ...FILE_MAP.db,
          base64: allSaved
            ? savedDb!.toString('base64')
            : toBase64(null, bundle!.db),
        },
        changeLog: {
          ...FILE_MAP.changeLog,
          base64: allSaved
            ? savedChangeLog!.toString('base64')
            : toBase64(null, bundle!.changeLog),
        },
      },
    });
  } catch (error: any) {
    console.error('Failed to load bundle:', error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
