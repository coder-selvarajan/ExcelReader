document.getElementById('input-excel').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, {
            type: 'binary'
        });

        const tabsContainer = document.getElementById('tabs');
        tabsContainer.innerHTML = ''; // Clear previous tabs
        const sheetsContainer = document.getElementById('sheets');
        sheetsContainer.innerHTML = ''; // Clear previous sheets

        workbook.SheetNames.forEach((sheetName, index) => {
            // Create tab
            const tab = document.createElement('div');
            tab.className = 'tab';
            tab.textContent = sheetName;
            tab.onclick = () => showSheet(index);
            tabsContainer.appendChild(tab);

            // Create sheet container
            const sheetContainer = document.createElement('div');
            sheetContainer.className = 'sheet-container';
            sheetContainer.style.display = 'none'; // Hide by default
            sheetsContainer.appendChild(sheetContainer);

            // Convert sheet to HTML and append
            const sheet = workbook.Sheets[sheetName];
            const sheetHTML = XLSX.utils.sheet_to_html(sheet, {id: "table" + index, editable: true});
            sheetContainer.innerHTML = sheetHTML;
        });

        // Show first sheet by default
        showSheet(0);
    };
    reader.readAsBinaryString(e.target.files[0]);
});

function showSheet(index) {
    const tabs = document.querySelectorAll('.tab');
    const sheets = document.querySelectorAll('.sheet-container');
    tabs.forEach(tab => tab.classList.remove('active'));
    sheets.forEach(sheet => sheet.style.display = 'none');

    tabs[index].classList.add('active');
    sheets[index].style.display = 'block';
}
