import * as XLSX from 'xlsx';

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
    Object.entries(days || {}).forEach(([day, meals]) => {
      let dayCalories = 0;
      let dayProtein = 0;

      if (Array.isArray(meals)) {
        meals.forEach(meal => {
          if (meal.ingredients && Array.isArray(meal.ingredients)) {
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
          if (meal.ingredients && Array.isArray(meal.ingredients)) {
            meal.ingredients.forEach((ing, index) => {
              dayData.push([
                index === 0 ? (meal.meal_name || 'Unnamed Meal') : '',
                ing.name || 'Unknown',
                ing.quantity || 0,
                ing.unit || '',
                ing.calories || 0,
                ing.protein || 0
              ]);
            });
            dayData.push(['', '', '', '', '', '']); // Empty row
          }
        });

        const dayWs = XLSX.utils.aoa_to_sheet(dayData);
        XLSX.utils.book_append_sheet(wb, dayWs, day);
      }
    });

    // Add summary sheet
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // Generate Excel file
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    // Set headers for direct download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="meal-plan.xlsx"');
    res.setHeader('Content-Length', buffer.length);

    // Send the file directly
    return res.send(buffer);

  } catch (error) {
    console.error('Error:', error);
    return res.status(500).json({ error: 'Failed to generate spreadsheet', details: error.message });
  }
}
