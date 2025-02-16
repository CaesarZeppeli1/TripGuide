const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;

// Middleware to parse form data
app.use(bodyParser.urlencoded({ extended: true }));

// Serve static files (HTML file in the 'public' folder)
app.use(express.static('public'));

// Route to handle form submission
app.post('/submit-form', (req, res) => {
    const formData = req.body;

    // Define the path to the Excel file
    const excelFilePath = path.join(__dirname, 'form_data.xlsx');

    // Initialize data array with header
    let data = [
        ['Name', 'Go Date', 'Leave Date', 'Email', 'Contact', 'Food Preference', 'Flavor', 'Places', 'Transport', 'Remarks']
    ];

    // Check if the Excel file exists
    if (fs.existsSync(excelFilePath)) {
        // Read the existing Excel file
        const workbook = xlsx.readFile(excelFilePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Convert existing sheet data to JSON format
        const existingData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Remove the header and append the existing data to the new data array
        existingData.forEach((row, index) => {
            if (index > 0) { // Skip the header
                data.push(row);
            }
        });
    }

    // Append new data
    data.push([
        formData.name,
        formData.dog,
        formData.dol,
        formData.email,
        formData.contact,
        formData.Select,
        formData.flavor,
        formData.place,
        formData.transportation,
        formData.Remarks
    ]);

    // Create a new workbook and add data
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Form Data');

    // Save the updated workbook to the Excel file
    xlsx.writeFile(workbook, excelFilePath);

    // Send response to user
    res.send('Form data saved to Excel.');
});

// Start the server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}/homePage.html`);
});
