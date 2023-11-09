const express = require('express');
const fs = require('fs');
const bodyParser = require('body-parser');
const app = express();
const cors = require('cors');

const Excel = require('exceljs');

app.use(cors());
app.use(bodyParser.json());

const daysOfWeek = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

// Check if the server can read staff_availability.json when it starts up
fs.readFile('staff_availability.json', 'utf8', (err, data) => {
    if (err) {
        console.error('Startup check: Error reading staff_availability.json:', err);
        return;
    }
    console.log('Startup check: Successfully read staff_availability.json');
    console.log('Startup check: First 100 characters of data:', data.substring(0, 100));  // Just a snippet to confirm
});

app.get('/staff-availability', (req, res) => {
    fs.readFile('staff_availability.json', 'utf8', (err, data) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            res.status(500).send('Error reading file');
            return;
        }
        console.log('Successfully read staff_availability.json during a GET request');
        res.send(data);
    });
});

app.post('/update-booked-dates', (req, res) => {
    console.log("Received request to update booked dates");
    const updatedData = req.body;
    console.log("Data received:", updatedData);  // Log the received data

    fs.writeFile('staff_availability.json', JSON.stringify(updatedData, null, 4), 'utf8', (err) => {
        if (err) {
            console.error('Error writing to staff_availability.json:', err);
            res.status(500).send('Error writing to file');
            return;
        }
        console.log('Successfully updated staff_availability.json');
        res.send({ success: true });
    });
});

app.post('/update-staff-availability', (req, res) => {
    const { staffName, updatedData } = req.body;

    if (!staffName || !updatedData) {
        return res.status(400).json({ success: false, message: "Missing required data" });
    }

    // Read the current JSON
    fs.readFile('staff_availability.json', 'utf8', (err, data) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            return res.status(500).send('Error reading file');
        }

        const currentData = JSON.parse(data);

        // Update the staff member's data
        currentData[staffName] = updatedData;

        // Write the updated data back to the JSON file
        fs.writeFile('staff_availability.json', JSON.stringify(currentData, null, 4), 'utf8', (err) => {

            if (err) {
                console.error('Error writing to staff_availability.json:', err);
                return res.status(500).send('Error writing to file');
            }

            res.json({ success: true, message: "Staff availability updated successfully!" });
        });
    });
});


// Endpoint to save the schedule as an Excel file
app.post('/save-schedule', async (req, res) => {
    console.log("SERVER STARTED SAVE ATTEMPT");
    console.log(req.body); // Log the entire body
    const { dates, scheduleData } = req.body;
    
    // Create a new workbook and add a sheet
    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('Schedule');

// Define the header row and apply styles
const headers = ['Shifts', ...dates]; // Include a 'Section' column before the dates
const headerRow = sheet.addRow(headers);

// Colours for schedule file styling
const lightBlueFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFADD8E6' } // Light blue
};

const lightGreenFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF90EE90' } // Light green
};

// Define the font style for Times New Roman
const timesNewRomanFont = {
  name: 'Times New Roman',
  size: 12 // You can specify the size as needed
};

// Apply styles to the header row cells here
headerRow.eachCell((cell) => {
  cell.font = { 
   ...timesNewRomanFont, // Spread the Times New Roman font settings
    bold: true, size: 13 };
  cell.fill = lightBlueFill; // Set the fill for header row cells
});

// Iterate through the properties of scheduleData to add rows to the sheet
Object.entries(scheduleData).forEach(([sectionName, days], index) => {
  // Start a new row with the section name
  const rowValues = [sectionName];
  
  
  // Push the values for each day into the rowValues array
  daysOfWeek.forEach(day => {
      rowValues.push(days[day] || ''); // If there's no entry for the day, push an empty string
  });
  
  // Add the row to the sheet and set the first cell to bold
  const row = sheet.addRow(rowValues);
  row.getCell(1).font = { bold: true, size: 13 };
  row.getCell(1).fill = lightGreenFill;
});




// Set the calculated width to each column after adding all rows
sheet.columns.forEach((column, index) => {
    // Add a buffer for padding, adjust the buffer size as needed
    const buffer = 5;

    // Filter out undefined or null values before calculating the length
    const columnValues = column.values.filter(val => val !== null && val !== undefined);

    const maxLength = columnValues.reduce((max, val) => {
        // If the value is a date, set a fixed length for formatted dates
        if (val instanceof Date) {
            return Math.max(max, 10); // Adjust this as per your date format
        }
        // Convert the value to string and calculate its length
        return Math.max(max, String(val).length);
    }, 0);

    // Apply buffer and set minimum width, if necessary
    column.width = maxLength + buffer > 10 ? maxLength + buffer : 10;
});

    // Write to a file
    try {
        const filePath = 'schedule.xlsx';
        await workbook.xlsx.writeFile(filePath);
        res.json({ success: true, message: 'Schedule saved successfully!', filePath });
    } catch (err) {
        console.error('Error writing to schedule.xlsx:', err);
        res.status(500).send('Error writing to file');
    }
});

app.get('/download-schedule', async (req, res) => {
    const filePath = 'schedule.xlsx';

    // Set the headers
    res.setHeader('Content-Disposition', 'attachment; filename=schedule.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    try {
        // Create a read stream and pipe it to the response
        const filestream = fs.createReadStream(filePath);
        filestream.pipe(res);
    } catch (err) {
        console.error('Error sending the file:', err);
        res.status(500).send('Error sending the file');
    }
});


app.get('/test-connection', (req, res) => {
  console.log("Test connection endpoint hit");
  res.json({ message: 'Connection successful!' });
});


const convertScheduleToCSV = (schedule) => {
    let csvArr = [];
    const daysOfWeek = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
    const slots = ["Lunch1", "Lunch2", "Bothams1", "Bothams2", "Bothams3", "Hole1", "Hole2" ];

    csvArr.push(['', ...daysOfWeek]);

    slots.forEach(slot => {
        let row = [];
        row.push(slot);

        daysOfWeek.forEach(day => {
            row.push(schedule[slot][day] || '');
        });

        csvArr.push(row);
    });

    return csvArr.map(e => e.join(",")).join("\n");
};


app.post('/update-max-shifts', (req, res) => {
    const { staffName, maxShifts } = req.body;

    fs.readFile('staff_availability.json', 'utf8', (err, data) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            res.status(500).send('Error reading file');
            return;
        }

        const currentData = JSON.parse(data);
        if (currentData[staffName]) {
            currentData[staffName].max_shifts = maxShifts;
        }

        fs.writeFile('staff_availability.json', JSON.stringify(currentData, null, 4), 'utf8', (err) => {
            if (err) {
                console.error('Error writing to staff_availability.json:', err);
                res.status(500).send('Error writing to file');
                return;
            }
            res.send({ success: true });
        });
    });
});
app.listen(3001, () => {
    console.log('Server started on http://localhost:3001');
});
