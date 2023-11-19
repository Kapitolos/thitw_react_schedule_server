const express = require('express');
const fs = require('fs');
const bodyParser = require('body-parser');
const app = express();
const cors = require('cors');
const router = express.Router();

const Excel = require('exceljs');

app.use(cors());
app.use(bodyParser.json());
app.use(router);

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

// Get restricted slots
router.get('/restricted-slots', (req, res) => {
    fs.readFile('restricted_slots.json', 'utf8', (err, data) => {
        if (err) {
            res.status(500).send('Error reading restricted slots');
            return;
        }
        res.json(JSON.parse(data));
    });
});

// Update restricted slots
router.post('/restricted-slots', (req, res) => {
    const newRestrictedSlots = req.body;
    fs.writeFile('restricted_slots.json', JSON.stringify(newRestrictedSlots, null, 4), 'utf8', (err) => {
        if (err) {
            res.status(500).send('Error updating restricted slots');
            return;
        }
        res.send({ success: true });
           res.json({ success: true, message: "Restricted slots updated successfully" });
    });
});

module.exports = router;

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


app.post('/add-staff', (req, res) => {
    const newStaffData = req.body;
    console.log("Received request to add new staff member");
    console.log("Data received:", newStaffData);  // Log the received data

    // Read the current staff availability data
    fs.readFile('staff_availability.json', 'utf8', (err, data) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            res.status(500).send('Error reading from file');
            return;
        }

        // Parse the current data and add the new staff member
        let staffAvailability = JSON.parse(data);
        const newStaffName = Object.keys(newStaffData)[0]; // Assuming newStaffData is an object with the staff name as the key

        // Check if staff member already exists to avoid duplicates
        if (staffAvailability[newStaffName]) {
            res.status(400).send('Staff member already exists');
            return;
        }

        // Add the new staff member's data
        staffAvailability[newStaffName] = newStaffData[newStaffName];

        // Write the updated data back to the JSON file
        fs.writeFile('staff_availability.json', JSON.stringify(staffAvailability, null, 4), 'utf8', (err) => {
            if (err) {
                console.error('Error writing to staff_availability.json:', err);
                res.status(500).send('Error writing to file');
                return;
            }
            console.log('Successfully added new staff member to staff_availability.json');
            res.send({ success: true, message: 'New staff member added' });
        });
    });
});

app.post('/remove-staff', (req, res) => {
    const { staffName } = req.body;

    // Read the current staff availability data
    fs.readFile('staff_availability.json', 'utf8', (err, data) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            res.status(500).send('Error reading from file');
            return;
        }

        let staffAvailability = JSON.parse(data);

        // Remove the staff member if they exist
        if (staffAvailability[staffName]) {
            delete staffAvailability[staffName];

            // Write the updated data back to the JSON file
            fs.writeFile('staff_availability.json', JSON.stringify(staffAvailability, null, 4), 'utf8', (err) => {
                if (err) {
                    console.error('Error writing to staff_availability.json:', err);
                    res.status(500).send('Error writing to file');
                    return;
                }
                console.log(`Successfully removed staff member: ${staffName}`);
                res.send({ success: true, message: `Staff member ${staffName} removed` });
            });
        } else {
            res.status(404).send({ success: false, message: `Staff member ${staffName} not found` });
        }
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

const yellowFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F5F580' } // Yellow
};

const orangeFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFBE4D' } // Orange
};

const redFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FD594D' } // Red
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

// Apply colors to section rows based on the section name
Object.entries(scheduleData).forEach(([sectionName, days], index) => {
    const rowValues = [sectionName];

 // Push the values for each day into the rowValues array
  daysOfWeek.forEach(day => {
      rowValues.push(days[day] || ''); // If there's no entry for the day, push an empty string
  });

    // Add the row to the sheet
    const row = sheet.addRow(rowValues);

    // Apply styles based on the sectionName
    let fill;
    if (sectionName.toLowerCase().includes('lunch')) {
        fill = yellowFill;
    } else if (sectionName.toLowerCase().includes('bothams')) {
        fill = orangeFill;
    } else if (sectionName.toLowerCase().includes('hole')) {
        fill = redFill;
    }

    // Set the font and fill for the first cell
    row.getCell(1).font = { bold: true, size: 13 };
    row.getCell(1).fill = fill;

    // Optionally, you could set the fill for the entire row:
    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
        cell.fill = fill;
    });
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
