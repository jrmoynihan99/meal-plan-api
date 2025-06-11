import { kv } from '@vercel/kv';

export default async function handler(req, res) {
  const { fileId } = req.query;

  if (!fileId) {
    return res.status(400).json({ error: 'File ID required' });
  }

  try {
    // Get file from KV storage
    const base64Buffer = await kv.get(fileId);
    
    if (!base64Buffer) {
      return res.status(404).json({ error: 'File not found or expired' });
    }

    // Convert back to buffer
    const buffer = Buffer.from(base64Buffer, 'base64');

    // Set headers to trigger download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="meal-plan.xlsx"');

    // Send the file
    return res.send(buffer);

  } catch (error) {
    console.error('Error downloading file:', error);
    return res.status(500).json({ error: 'Failed to download file' });
  }
}
