const fs = require('fs');
const sched = JSON.parse(fs.readFileSync('trip_schedule_data.json', 'utf8'));
const d7 = sched.find(x => x.dayNumber === 7);
const d8 = sched.find(x => x.dayNumber === 8);
console.log('Day 7:');
d7.steps.forEach((s, i) => console.log('  Step ' + (i+1) + ': [' + s.time + '] ' + s.item));
console.log('Day 8:');
d8.steps.forEach((s, i) => console.log('  Step ' + (i+1) + ': [' + s.time + '] ' + s.item));
