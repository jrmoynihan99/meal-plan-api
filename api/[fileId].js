import { readFileSync, existsSync } from 'fs';

export default async function handler(req, res) {
  const { fileId } = req.query;

  if (!fileId) {
    return res.status(400).json({ error: 'File ID required' });
  }

  const filePath = `/tmp/${fileId}.xlsx`;

  try {
    // Check if file exists
    if (!existsSync(filePath)) {
      return res.status(404).json({ error: 'File not found or expired' });
    }

    // Read the file
    const buffer = readFileSync(filePath);

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
