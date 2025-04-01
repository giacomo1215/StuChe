document.addEventListener('DOMContentLoaded', () => {                   // Ensure the DOM is fully loaded
    const resultsPanel = document.getElementById('results');            // Get the results panel element
    const resultCard = document.getElementById('resultsCard');          // Get the results card element
    const excelFileInput = document.getElementById('excelFile');        // Get the file input element
    const errorMessageBody = document.getElementById('errorMessage');   // Get the error message body element
    const errorMessage = document.getElementById('errorAlert');         // Get the error message alert element

    try {
        excelFileInput.addEventListener('change', (event) => {
            const file = event.target.files[0];                             // Get the selected file
            const reader = new FileReader();                                // Create a new FileReader instance

            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);               // Read the file data
                const workbook = XLSX.read(data, { type: 'array' });        // Parse the Excel file
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];  // Get the first worksheet
                const students = XLSX.utils.sheet_to_json(worksheet);       // Convert the worksheet to JSON

                const hasApptIssue = (value) =>                             // Check if there's issues
                    value === undefined ||                                  // Undefined value
                    value === "" ||                                         // Empty string
                    value?.includes("No") ||                                // Contains "No"
                    value?.includes("-");                                   // Contains a dash

                // Filter students with appointment issues
                const flaggedStudents = students.filter(student =>
                    ['1 Appt', '2 Appt', '3 Appt'].some(appt => hasApptIssue(student[appt]))
                );

                // Determine specific issues for flagged students
                flaggedStudents.forEach(student => {                                // Iterate through flagged students
                    const appointments = ['1 Appt', '2 Appt', '3 Appt'];            // Appointment keys
                    
                    for (const appt of appointments) {                              // Iterate through appts
                        const value = student[appt];                                // Get the appointment value
                        const [apptNumber] = appt.split(' ');                       // Extract appointment number

                        switch (true) {
                            case value?.includes("No") || value === undefined:       // Check for "No" or undefined
                                student.Issue = `Missing ${apptNumber} appointment`; // Set issue
                                break;
                            case value === "":                                       // Check for empty string
                                student.Issue = `Empty ${apptNumber} appointment`;   // Set issue
                                break;
                            case value?.includes("-"):                               // Check for dash
                                student.Issue = "Booked but not attended";           // Set issue
                                break;
                            default:                                                 // Default case
                                student.Issue = "Unknown issue";                     // Set issue
                        }
                    }
                });

                // Build the result table in case there are flagged students
                const html = flaggedStudents.length > 0
                ? flaggedStudents.map(student => `
                        <tr>
                            <td>${student.SID}</td>
                            <td>${student["First Name"]}</td>
                            <td>${student["Last Name "]}</td>
                            <td>${student.Issue}</td>
                        </tr>
                    `).join('')
                : "<tr><td colspan='4'>No flagged students found.</td></tr>";

                resultsPanel.innerHTML = html;                                      // Set the table header
                resultCard.style.display = "block";                                 // Show the results card
            };

            reader.readAsArrayBuffer(file);
        });
    } catch (error) {
        console.error(error);
        errorMessageBody.innerHTML = error.message;                         // Display the error message
        errorMessage.classList.remove('d-none');                            // Show the error alert
        resultCard.classList.add('d-none');                                 // Hide the results card
    }
});