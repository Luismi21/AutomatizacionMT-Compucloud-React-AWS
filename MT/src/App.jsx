import React, { useState } from 'react';
// Importa tus estilos
import './index.css';

// ¡Esta URL ya está correcta!
const API_GATEWAY_URL = 'https://s9yurg9hj8.execute-api.us-east-1.amazonaws.com/generate';

function App() {
  const [selectedFile, setSelectedFile] = useState(null);
  const [fileName, setFileName] = useState('Ningún archivo seleccionado');
  const [htmlPreview, setHtmlPreview] = useState(null);
  const [downloadUrl, setDownloadUrl] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file) {
      if (file.type !== 'application/json') {
          setError('Error: El archivo debe ser de tipo .json');
          setSelectedFile(null);
          setFileName('Ningún archivo seleccionado');
      } else {
          setSelectedFile(file);
          setFileName(file.name);
          setError(null);
      }
    } else {
      setSelectedFile(null);
      setFileName('Ningún archivo seleccionado');
    }
  };

  // --- handleSubmit (VERSIÓN REAL, SIN SIMULACIÓN) ---
  const handleSubmit = async (event) => {
    event.preventDefault(); 
    
    if (!selectedFile) {
      setError('Por favor, selecciona un archivo .json');
      return;
    }

    setIsLoading(true);
    setError(null);
    setHtmlPreview(null);
    setDownloadUrl(null);

    try {
      const fileReader = new FileReader();
      // Lee el archivo como un string Base64
      fileReader.readAsDataURL(selectedFile); 
      
      // Se activa CUANDO la lectura del archivo se completa
      fileReader.onload = async (e) => {
        try {
          // 1. Quita el prefijo 'data:application/json;base64,'
          const base64Content = e.target.result.split(',')[1];
          
          // 2. Llama a la API Gateway
          const response = await fetch(API_GATEWAY_URL, {
            method: 'POST',
            headers: { 
              'Content-Type': 'application/json' 
            },
            body: base64Content // Envía el string base64
          });
  
          const data = await response.json();
  
          if (!response.ok) {
            // Si la Lambda devuelve un error (statusCode 500)
            throw new Error(data.error || 'Ocurrió un error en el servidor');
          }
  
          // 3. ¡Éxito! Actualiza el estado
          setHtmlPreview(data.html_preview);
          setDownloadUrl(data.download_url);

        } catch (apiError) {
           // Error durante el 'fetch' o si la respuesta no es 'ok'
          console.error('Error de API:', apiError);
          setError(apiError.message);
        } finally {
          // Esto se ejecuta después del try/catch del 'fetch'
          setIsLoading(false); 
        }
      };

      // Se activa si la LECTURA del archivo falla
      fileReader.onerror = (err) => {
        console.error('Error leyendo archivo:', err);
        setError('Error al leer el archivo local.');
        setIsLoading(false);
      };

    } catch (generalError) {
      // Error si algo falla antes (ej: new FileReader())
      console.error('Error general:', generalError);
      setError(generalError.message);
      setIsLoading(false);
    }
  };
  // --- FIN DE handleSubmit ---

  return (
    <> 
      <div className="card">
        <div className="card-header">
          <h1>
            <svg style={{ verticalAlign: 'middle', width: '28px', height: '28px', marginRight: '10px' }} xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zM18 20H6V4h7v5h5v11z" /></svg>
            Generador de Memoria Técnica
          </h1>
        </div>
        <div className="card-body">
          <p>Sube tu archivo <code>estado_infraestructura.json</code> para generar el documento de Word automáticamente.</p>
          
          {error && (
            <div className="flash-error">{error}</div>
          )}

          <form onSubmit={handleSubmit}>
            <label htmlFor="file-upload" className="custom-file-upload">
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path d="M9 16h6v-6h4l-7-7-7 7h4v6zm-4 2h14v2H5v-2z" /></svg>
              Seleccionar Archivo
            </label>
            <input 
              id="file-upload" 
              type="file" 
              name="file" 
              accept=".json" 
              onChange={handleFileChange} 
            />
            <span id="file-name" className="file-name-display">{fileName}</span>
            
            <button type="submit" className="submit-btn" disabled={isLoading}>
              {isLoading ? 'Generando...' : 'Generar Documento'}
            </button>
          </form>
        </div>
      </div>

      {htmlPreview && (
        <div className="card preview-card">
          <div className="card-header">
            <h2>Vista Previa</h2>
          </div>
          <div className="card-body">
            <div 
              className="document-preview"
              dangerouslySetInnerHTML={{ __html: htmlPreview }}
            />
          </div>
          <div className="download-container">
            <a href={downloadUrl} className="submit-btn download-btn" download>
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="currentColor"><path d="M5 20h14v-2H5v2zM19 9h-4V3H9v6H5l7 7 7-7z" /></svg>
              Descargar .DOCX
            </a>
          </div>
        </div>
      )}
    </>
  );
}

export default App;

