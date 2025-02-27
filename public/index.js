const form = document.getElementById('uploadForm');
form.addEventListener('submit', async (e) => {
    e.preventDefault()

    const fileInput = document.getElementById('excelFile');
    const formData = new FormData();

    formData.append('excelFile', fileInput.files[0])

    console.log(formData)

    try {
        const response = await fetch('/upload', {
          method: 'POST',
          body: formData
        });
        const result = await response.json();
        console.log(result);
      } catch (error) {
        console.error('Error uploading file:', error);
      }
})

const downloadBtn = document.getElementById('getUpdatedFile');
downloadBtn.addEventListener('click', async (e) => {
  e.preventDefault()

  try {
    const response = await fetch('/get-new-file');
    if (!response.ok) throw new Error('Failed to download file');

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'new-file.xlsx'; // Ensure correct filename
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  } catch (error) {
    console.error('Error uploading file:', error);
  }
})
