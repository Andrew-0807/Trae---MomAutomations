document.addEventListener('DOMContentLoaded', () => {
  // Get references to DOM elements
  const processBtn = document.getElementById('processBtn');
  const fileInput = document.getElementById('fileInput');
  
  if (!processBtn || !fileInput) {
    console.error('Required elements not found. Make sure your HTML has elements with IDs "processBtn" and "fileInput"');
    return;
  }

  processBtn.addEventListener('click', () => {
    const files = fileInput.files;
    const processType = document.querySelector('input[name="process_type"]:checked').value;
    console.log('Clicked');
    if (!files.length) {
      alert('Please select a file.');
      return;
    }
  
    const formData = new FormData();
    // Append all files to support multiple file uploads
    for (let i = 0; i < files.length; i++) {
        formData.append('file', files[i]);
    }
    formData.append('process_type', processType);
  
    fetch('/process', {
      method: 'POST',
      body: formData
    })
    .then(response => {
      if (!response.ok) throw new Error('Network response was not OK');
      
      // Extract filename from Content-Disposition header
      const disposition = response.headers.get('Content-Disposition');
      let filename = files[0].name; // fallback to original filename
      
      if (disposition) {
        // More robust filename extraction
        const filenameMatch = disposition.match(/filename\*?=['"]?(?:UTF-\d['"]*)?([^;\r\n"']*)['"]?;?/i);
        if (filenameMatch && filenameMatch[1]) {
          filename = filenameMatch[1];
          console.log('Extracted filename:', filename);
        }
      }
      
      return response.blob().then(blob => ({ blob, filename }));
    })
    .then(({ blob, filename }) => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
      }, 100);
    })
    .catch(err => {
      console.error('Error processing file:', err);
      alert('Error processing file: ' + err.message);
    });
  });
});

// Remove this line as it's causing an error and appears to be leftover code
// const match = disposition.match(/filename="?(.+?)"?($|;)/);
  