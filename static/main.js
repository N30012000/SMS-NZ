let tmpdir = null;
document.getElementById('uploadBtn').onclick = async () => {
  const files = document.getElementById('fileinput').files;
  if (!files.length) { alert('Select files'); return; }
  let form = new FormData();
  for (let f of files) form.append('files[]', f);
  const res = await fetch('/upload', {method:'POST', body: form});
  const j = await res.json();
  if (j.tmpdir) {
    tmpdir = j.tmpdir;
    alert('Uploaded ' + j.count + ' files. Ready to extract.');
  } else {
    alert('Upload failed');
  }
};

document.getElementById('extractBtn').onclick = async () => {
  if (!tmpdir) { alert('Upload first'); return; }
  const res = await fetch('/extract', {
    method:'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({tmpdir: tmpdir})
  });
  const j = await res.json();
  if (j.excel_path) {
    const dl = document.getElementById('downloadLink');
    dl.style.display = 'inline';
    dl.href = '/download?path=' + encodeURIComponent(j.excel_path);
    dl.innerText = 'Download Extracted Excel';
  } else {
    alert('Extraction failed: ' + (j.error || 'unknown'));
  }
};

document.getElementById('genDash').onclick = async () => {
  const monthVal = document.getElementById('monthpicker').value;
  const file = document.getElementById('excelfile').files[0];
  if (!monthVal || !file) { alert('Select month and Excel file'); return; }
  const [year,month] = monthVal.split('-');
  let form = new FormData();
  form.append('file', file);
  form.append('month', month);
  form.append('year', year);
  const res = await fetch('/dashboard', {method:'POST', body: form});
  const j = await res.json();
  if (j.html_preview) {
    document.getElementById('dashResult').innerHTML = `<a href="${j.html_preview}">Open Dashboard Preview</a><br><a href="/download?path=${encodeURIComponent(j.excel)}">Download Dashboard Excel</a><br><a href="/download?path=${encodeURIComponent(j.pdf)}">Download Dashboard PDF</a>`;
  } else {
    alert('Dashboard generation failed');
  }
};
