import { NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

export async function POST(req: Request) {
  try {
    const { files } = await req.json();
    if (!files) {
      return NextResponse.json({ success: false, error: 'No files provided' }, { status: 400 });
    }

    // 連携用のローカルフォルダ (.env.local から取得。指定がなければプロジェクト直下のフォルダを使用)
    const driveDir = process.env.GOOGLE_DRIVE_MOCK_DIR || path.join(process.cwd(), 'CloudDriveSync');
    
    // フォルダが存在しなければ作成
    if (!fs.existsSync(driveDir)) {
      fs.mkdirSync(driveDir, { recursive: true });
    }

    // 各ファイルをBase64からデコードして保存
    const fileMap: Record<string, string> = {
      screen: '画面設計書_SCR001.xlsx',
      api: 'API仕様書_users.xlsx',
      db: 'テーブル定義書_users.xlsx',
      changeLog: '変更管理簿_2026.xlsx',
    };

    const savedFiles: string[] = [];

    for (const [key, base64Data] of Object.entries(files)) {
      if (typeof base64Data === 'string' && fileMap[key]) {
        const filePath = path.join(driveDir, fileMap[key]);
        const buffer = Buffer.from(base64Data, 'base64');
        fs.writeFileSync(filePath, buffer);
        savedFiles.push(fileMap[key]);
      }
    }

    return NextResponse.json({
      success: true,
      message: 'Successfully synced to local Drive folder',
      savedFiles,
      driveDir
    });
  } catch (error: any) {
    console.error('Drive sync error:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
