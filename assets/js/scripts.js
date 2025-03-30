document.addEventListener('DOMContentLoaded', () => {                   // Ensure the DOM is fully loaded
    const resultsPanel = document.getElementById('results');            // Get the results panel element
    const resultCard = document.getElementById('resultsCard');          // Get the results card element
    const excelFileInput = document.getElementById('excelFile');        // Get the file input element

    excelFileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];                             // Get the selected file
        const reader = new FileReader();                                // Create a new FileReader instance

        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);               // Read the file data
            const workbook = XLSX.read(data, { type: 'array' });        // Parse the Excel file
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];  // Get the first worksheet
            const students = XLSX.utils.sheet_to_json(worksheet);       // Convert the worksheet to JSON

            const flaggedStudents = students.filter(student =>          // Filter students based on criteria
                !student["1 Appt"] ||                                   // Check if "1 Appt" is empty
                student["2 Appt"]?.includes("-") ||                     // Check if "2 Appt" contains "-"
                student["3 Appt"]?.includes("-")                        // Check if "3 Appt" contains "-"
            );

            let html = "";
            flaggedStudents.forEach(student => {                        // Iterate over flagged students
                html += "<tr>";                                         // Start a new table row
                html += `<td>${student["SID"]}</td>`;                   // Add student ID
                html += `<td>${student["First Name"]}</td>`;            // Add student name
                html += `<td>${student["Last Name "]}</td>`;            // Add student name
                html += `<td>${student["1 Appt"] || "N/A"}</td>`;       // Add "1 Appt" value
                html += "</tr>";                                        // Close the table row
            });
            if (flaggedStudents.length === 0) {                         // If no students are flagged
                html = "<tr><td colspan='4'>No flagged students found.</td></tr>"; // Show a message
            }
            resultsPanel.innerHTML = html;                              // Update the results panel with the HTML
            resultCard.style.display = "block";                         // Show the results card
        };

        reader.readAsArrayBuffer(file);
    });
});