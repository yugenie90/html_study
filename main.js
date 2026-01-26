document.getElementById('report-form').addEventListener('submit', function(event) {
    event.preventDefault();
    const reportInput = document.getElementById('report-input').value;
    const reportOutput = document.getElementById('report-output');

    if (reportInput.trim() === '') {
        alert('보고서를 생성하려면 텍스트를 입력하세요.');
        return;
    }

    reportOutput.innerHTML = '<p>보고서를 생성하는 중...</p>';

    // Simulate a delay for report generation
    setTimeout(() => {
        reportOutput.innerHTML = `<p>입력하신 내용 기반의 샘플 보고서입니다: "${reportInput}"</p>`;
    }, 2000);
});