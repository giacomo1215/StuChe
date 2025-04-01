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

                const hasApptIssue = (value) => 
                    value === undefined ||
                    value === "" ||
                    value?.includes("No") ||
                    value?.includes("-");

                const flaggedStudents = students.filter(student =>
                    ['1 Appt', '2 Appt', '3 Appt'].some(appt => hasApptIssue(student[appt]))
                );

                // Determine specific issues for flagged students
                flaggedStudents.forEach(student => {
                    const appointments = ['1 Appt', '2 Appt', '3 Appt'];
                    
                    for (const appt of appointments) {
                        const value = student[appt];
                        const [apptNumber] = appt.split(' ');

                        if (value?.includes("No") || value === undefined) {
                            student.Issue = `Missing ${apptNumber} appointment`;
                            break;
                        } else if (value === "") {
                            student.Issue = `Empty ${apptNumber} appointment`;
                            break;
                        } else if (value?.includes("-")) {
                            student.Issue = "Booked but not attended";
                            break;
                        }
                    }
                });

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

                resultsPanel.innerHTML = html;
                resultCard.style.display = "block";            // Show the results card
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