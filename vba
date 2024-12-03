const data = [
  { variableA: 10, variableB: 20, category: ['Category1', 'Category3'] },
  { variableA: 15, variableB: 25, category: ['Category2'] },
  { variableA: 20, variableB: 30, category: ['Category1', 'Category2'] },
  { variableA: 25, variableB: 35, category: ['Category3'] },
  { variableA: 30, variableB: 40, category: ['Category2', 'Category3'] },
];

// Grouping by categories
const groupedData = data.reduce((result, item) => {
  item.category.forEach((cat) => {
    // If the category doesn't exist, initialize it as an empty array
    if (!result[cat]) {
      result[cat] = [];
    }
    // Add the current item to the corresponding category
    result[cat].push(item);
  });
  return result;
}, {});

console.log(groupedData);
