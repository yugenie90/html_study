document.getElementById('report-form').addEventListener('submit', function(event) {
    event.preventDefault();
    const reportInput = document.getElementById('report-input').value;
    const reportOutput = document.getElementById('report-output');

    if (reportInput.trim() === '') {
        alert('Please enter some text to generate a report.');
        return;
    }

    reportOutput.innerHTML = '<p>Generating report...</p>';

    // Simulate a delay for report generation
    setTimeout(() => {
        reportOutput.innerHTML = `<p>This is a sample report based on your input: "${reportInput}"</p>`;
    }, 2000);
});
