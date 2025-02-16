import React, { useState } from 'react';

const TemplateImport = ({ onTemplateLoaded }) => {
  const [fileName, setFileName] = useState(null);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      setFileName(file.name);
      const reader = new FileReader();
      reader.onload = (event) => {
        const arrayBuffer = event.target.result;
        onTemplateLoaded(arrayBuffer);
      };
      reader.readAsArrayBuffer(file);
    }
  };

  return (
    <div style={{ marginTop: '1rem' }}>
      <label htmlFor="template-upload">Template hochladen:</label>
      <input
        id="template-upload"
        type="file"
        accept=".docx"
        onChange={handleFileChange}
        style={{ marginLeft: '0.5rem' }}
      />
      {fileName && <p>Hochgeladen: {fileName}</p>}
    </div>
  );
};

export default TemplateImport;
