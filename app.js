function toFormUrlEncoded(str) {
    return encodeURIComponent(str).replace(/%20/g, '+');
  }

  function generateUrl() {
    const base = document.getElementById('baseInput').value.trim();
    const from = document.getElementById('fromInput').value.trim();
    const to = document.getElementById('toInput').value.trim();

    if (from === '') {
      document.getElementById('result').textContent = 'Mohon isi nama pengantin!';
      return;
    }

    const encodedFrom = toFormUrlEncoded(from);
    const finalUrl = to
      ? `${base}?from=${encodedFrom}&to=${toFormUrlEncoded(to)}`
      : `${base}?from=${encodedFrom}`;

    document.getElementById('result').textContent = finalUrl;
  }

  function handleExcel() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];

    if (!file) {
      alert("Silakan unggah file Excel terlebih dahulu.");
      return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      const base = document.getElementById('baseInput').value.trim();
      let from = document.getElementById('fromInput').value.trim();
      if (!from && jsonData.length > 0 && jsonData[0].from) {
        from = jsonData[0].from;
      }

      if (!from) {
        document.getElementById('bulkResult').textContent = 'Mohon isi nama pengantin di input atau Excel.';
        return;
      }

      const encodedFrom = toFormUrlEncoded(from);
      let output = '';
      jsonData.forEach((row, index) => {
        const toName = row.to ? toFormUrlEncoded(row.to) : '';
        const url = toName
          ? `${base}?from=${encodedFrom}&to=${toName}`
          : `${base}?from=${encodedFrom}`;
        output += `${index + 1}. ${decodeURIComponent(url)}\n`;
      });

      document.getElementById('bulkResult').textContent = output;
    };

    reader.readAsArrayBuffer(file);
  }