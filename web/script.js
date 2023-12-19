async function fetchDataAndModifyTable() {
    try {
      const response = await fetch('http://127.0.0.1:3000/api/students');
      const data = await response.json();
      const tableBody = document.getElementById('tableBody');
      let tableRows = '';
  
      data.forEach((student, index) => {
        const changeClass = student.difference > 0 ? 'positive' : student.difference < 0 ? 'negative' : 'neutral';
        const arrow = student.difference > 0 ? 'fa-arrow-up' : student.difference < 0 ? 'fa-arrow-down' : '';
  
        const row = `
          <tr>
            <td>${index + 1}<span class="change ${changeClass}"><i class="fas ${arrow}"></i> ${Math.abs(student.difference)}</span></td>
            <td>${student.studentName}</td>
            <td>${student.average}</td>
          </tr>
        `;
  
        tableRows += row;
      });
  
      tableBody.innerHTML = tableRows;
    } catch (error) {
      console.error('Error fetching data:', error);
    }
  }
  
  fetchDataAndModifyTable();
  