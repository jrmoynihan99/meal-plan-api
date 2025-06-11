import * as XLSX from 'xlsx';
import { v4 as uuidv4 } from 'uuid';
import { kv } from '@vercel/kv';

export default async function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { calorie_target, protein_target, days } = req.body;

    // Create workbook
    const wb = XLSX.utils.book_new();

    // Summary sheet data
    const summaryData = [
      ['Daily Calorie Target', calorie_target || 'Not specified'],
      ['Daily Protein Target', protein_target || 'Not specified'],
      [''],
      ['Day', 'Total Calories', 'Total Protein', '# of Meals']
    ];

    // Process each day
    Object.entries(days).forEach(([day, meals]) => {
      let dayCalories = 0;
      let dayProtein = 0;

      meals.forEach(meal => {
        if (meal.ingredients) {
          meal.ingredients.forEach(ing => {
            dayCalories += ing.calories || 0;
            dayProtein += ing.protein || 0;
          });
        }
      });

      summaryData.push([day, dayCalories, dayProtein, meals.length]);

      // Create detailed sheet for each day
      const dayData = [
        ['Meal Name', 'Ingredient', 'Quantity', 'Unit', 'Calories', 'Protein']
      ];

      meals.forEach(meal => {
        if (meal.ingredients) {
          meal.ingredients.forEach((ing, index) => {
            dayData.push([
              index === 0 ? meal.meal_name : '',
              ing.name,
              ing.quantity,
              ing.unit,
              ing.calories,
              ing.protein
            ]);
          });
          dayData.push(['', '', '', '', '', '']); // Empty row
        }
      });

      const dayWs = XLSX.utils.aoa_to_sheet(dayData);
      XLSX.utils.book_append_sheet(wb, dayWs, day);
    });

    // Add summary sheet
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // Generate Excel file
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    
    // Convert buffer to base64 for storage
    const base64Buffer = buffer.toString('base64');

    // Store in Vercel KV with expiration (1 hour = 3600 seconds)
    const fileId = uuidv4();
    await kv.set(fileId, base64Buffer, { ex: 3600 });

    // Return download link
    return res.json({
      success: true,
      download_url: `https://meal-plan-api-one.vercel.app/api/download/${fileId}`,
      filename: 'meal-plan.xlsx'
    });

  } catch (error) {
    console.error('Error:', error);
    return res.status(500).json({ error: 'Failed to generate spreadsheet' });
  }
}
