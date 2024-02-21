
const express = require('express');
const fs = require('fs');
const bodyParser = require('body-parser');
const AWS = require('aws-sdk');
const app = express();
const cors = require('cors');
const router = express.Router();
const stream = require('stream');
const Excel = require('exceljs');

app.use(cors());
app.use(bodyParser.json());
app.use(router);

const daysOfWeek = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

// Set up AWS S3
const s3 = new AWS.S3({
  accessKeyId: process.env.AWS_ACCESS_KEY_ID,
  secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
  region: process.env.AWS_REGION
});

const S3_BUCKET = process.env.AWS_BUCKET_NAME;
console.log("S3 Bucket Name:", process.env.AWS_BUCKET_NAME);

// Helper function to get JSON from S3
function getJsonFromS3(key, callback) {
  s3.getObject({ Bucket: S3_BUCKET, Key: key }, (err, data) => {
    if (err) return callback(err);
    try {
      const json = JSON.parse(data.Body.toString());
      callback(null, json);
    } catch (e) {
      callback(e);
    }
  });
}

// Helper function to upload JSON to S3
function uploadJsonToS3(key, json, callback) {
  const body = Buffer.from(JSON.stringify(json, null, 2));
  s3.putObject({ Bucket: S3_BUCKET, Key: key, Body: body }, callback);
}

app.get('/staff-availability', (req, res) => {
  getJsonFromS3('staff_availability.json', (err, data) => {
    if (err) {
      console.error('Error reading staff_availability.json:', err);
      return res.status(500).send('Error reading file');
    }
    res.json(data);
  });
});

app.post('/update-booked-dates', (req, res) => {
    console.log("Received request to update booked dates");
    const updatedData = req.body;
    console.log("Data received:", updatedData);  // Log the received data

    const key = 'staff_availability.json';

    uploadJsonToS3(key, updatedData, (err) => {
        if (err) {
            console.error('Error writing to staff_availability.json:', err);
            res.status(500).send('Error writing to file');
            return;
        }
        console.log('Successfully updated staff_availability.json');
        res.json({ success: true, message: "Booked dates updated successfully" });
    });
});

router.post('/update-staff-availability', (req, res) => {
  const { staffName, updatedData } = req.body;

  if (!staffName || !updatedData) {
    return res.status(400).json({ success: false, message: "Missing required data" });
  }

  getJsonFromS3('staff_availability.json', (err, currentData) => {
    if (err) {
      console.error('Error reading staff_availability.json:', err);
      return res.status(500).send('Error reading file');
    }

    // Update the staff member's data
    currentData[staffName] = updatedData;

    // Write the updated data back to S3
    uploadJsonToS3('staff_availability.json', currentData, (err) => {
      if (err) {
        console.error('Error writing to staff_availability.json:', err);
        return res.status(500).send('Error writing to file');
      }
      res.json({ success: true, message: "Staff availability updated successfully!" });
    });
  });
});

// Update restricted slots
router.post('/restricted-slots', (req, res) => {
    const newRestrictedSlots = req.body;
    const key = 'restricted_slots.json';

    uploadJsonToS3(key, newRestrictedSlots, (err) => {
        if (err) {
            console.error('Error updating restricted slots:', err);
            res.status(500).send('Error updating restricted slots');
            return;
        }
        res.json({ success: true, message: "Restricted slots updated successfully" });
    });
});

// Fetch restricted slots
app.get('/restricted-slots', (req, res) => {
    const key = 'restricted_slots.json'; // Ensure this matches your S3 file key

    s3.getObject({ Bucket: S3_BUCKET, Key: key }, (err, data) => {
        if (err) {
            console.error('Error fetching restricted slots:', err);
            return res.status(500).send('Error fetching restricted slots');
        }
        try {
            const slotsData = JSON.parse(data.Body.toString());
            res.json(slotsData);
        } catch (parseError) {
            console.error('Error parsing restricted slots data:', parseError);
            res.status(500).send('Error parsing restricted slots data');
        }
    });
});


app.post('/add-staff', (req, res) => {
    console.log("Received request to add new staff member");
    const newStaffData = req.body;
    console.log("Data received:", newStaffData);  // Log the received data

    getJsonFromS3('staff_availability.json', (err, staffAvailability) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            res.status(500).send('Error reading from file');
            return;
        }

        const newStaffName = Object.keys(newStaffData)[0];
        if (staffAvailability[newStaffName]) {
            res.status(400).send('Staff member already exists');
            return;
        }

        staffAvailability[newStaffName] = newStaffData[newStaffName];
        uploadJsonToS3('staff_availability.json', staffAvailability, (err) => {
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

    getJsonFromS3('staff_availability.json', (err, staffAvailability) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            res.status(500).send('Error reading from file');
            return;
        }

        if (staffAvailability[staffName]) {
            delete staffAvailability[staffName];
            uploadJsonToS3('staff_availability.json', staffAvailability, (err) => {
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

app.post('/update-max-shifts', (req, res) => {
    const { staffName, maxShifts } = req.body;

    getJsonFromS3('staff_availability.json', (err, currentData) => {
        if (err) {
            console.error('Error reading staff_availability.json:', err);
            res.status(500).send('Error reading file');
            return;
        }

        if (currentData[staffName]) {
            currentData[staffName].max_shifts = maxShifts;
            uploadJsonToS3('staff_availability.json', currentData, (err) => {
                if (err) {
                    console.error('Error writing to staff_availability.json:', err);
                    res.status(500).send('Error writing to file');
                    return;
                }
                res.send({ success: true, message: "Max shifts updated successfully" });
            });
        } else {
            res.status(404).send({ success: false, message: "Staff member not found" });
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

    // Instead of writing to a file, write to a stream
    const passThrough = new stream.PassThrough();
    workbook.xlsx.write(passThrough)
        .then(() => {
            passThrough.end();
        })
        .catch(err => {
            console.error('Error writing Excel stream:', err);
            res.status(500).send('Error generating file');
            return;
        });

            // Listen for errors on the stream to catch any errors that occur after initiating the upload
    passThrough.on('error', (streamError) => {
        console.error('Stream encountered an error:', streamError);
        res.status(500).send('Error in stream processing');
    });

          // Upload the stream to S3
    const filePath = 'schedule.xlsx';
    s3.upload({
        Bucket: S3_BUCKET,
        Key: filePath,
        Body: passThrough,
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }, (err, data) => {
        if (err) {
            console.error('Error uploading to S3:', err);
            res.status(500).send('Error uploading file');
            return;
        }
        console.log('File uploaded successfully:', data.Location);
        res.json({ success: true, message: 'Schedule saved successfully!', filePath: data.Location });
    });
});

app.get('/download-schedule', async (req, res) => {
    // Extract the date from the query parameter
    const date = req.query.date;
    if (!date) {
        return res.status(400).send('Date parameter is required');
    }

    const filename = `Schedule_${date}.xlsx`;
    const key = 'schedule.xlsx'; // The S3 key for the stored schedule file

    // S3 getObject parameters
    const params = {
        Bucket: S3_BUCKET,
        Key: key,
    };

    // Try to get the object from S3
    const fileStream = s3.getObject(params).createReadStream();
    fileStream.on('error', function(error) {
        console.error('Error streaming file from S3:', error);
        return res.status(500).send('Error streaming file');
    });

    // Set headers to suggest the filename and indicate the content type
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Stream the file from S3 to the client
    fileStream.pipe(res);
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


// The rest of the endpoints would be modified in a similar way...
// Process.env.PORT is provided by Heroku, 3001 is used for local development
const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`Server started on http://localhost:${PORT}`);
});