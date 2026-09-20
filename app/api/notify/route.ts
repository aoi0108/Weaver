// app/api/notify/route.ts
import { NextResponse } from 'next/server';

export async function POST(req: Request) {
  try {
    const { ticketId, summary, impact, author } = await req.json();
    const webhookUrl = process.env.SLACK_WEBHOOK_URL;

    if (!webhookUrl) {
      console.warn('SLACK_WEBHOOK_URL is not set.');
      return NextResponse.json({ success: false, message: 'No webhook URL' });
    }

    // Slack Block Kit で見栄えの良いメッセージを構築
    const payload = {
      text: `📢 【仕様変更確定】設計書が更新されました`,
      blocks: [
        {
          type: 'header',
          text: {
            type: 'plain_text',
            text: `🚀 【仕様変更確定】設計書が更新されました`,
            emoji: true,
          },
        },
        {
          type: 'section',
          fields: [
            ...(ticketId ? [{
              type: 'mrkdwn',
              text: `*変更チケット:*\n${ticketId}`,
            }] : []),
            {
              type: 'mrkdwn',
              text: `*承認者:*\n${author || 'プロダクトマネージャー'}`,
            },
          ],
        },
        {
          type: 'section',
          text: {
            type: 'mrkdwn',
            text: `*変更概要:*\n${summary}`,
          },
        },
        {
          type: 'section',
          text: {
            type: 'mrkdwn',
            text: `*影響範囲:*\n\`${impact}\``,
          },
        },
        {
          type: 'context',
          elements: [
            {
              type: 'mrkdwn',
              text: ' *Excel設計書直接更新完了*: 画面・API・DB・変更管理簿の全4シートがAI自律整合済み。\n📁 *Google Drive / Box* の最新設計書（.xlsx）が自動同期されました。ベンダー側担当者はご確認ください。',
            },
          ],
        },
        {
          type: 'divider',
        },
      ],
    };

    const res = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      throw new Error(`Slack API responded with status ${res.status}`);
    }

    return NextResponse.json({ success: true });
  } catch (error: any) {
    console.error('Slack notification error:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}