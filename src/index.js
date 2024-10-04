import React, { useState } from "react";
import { Form, Button, Container, Spinner, Alert } from "react-bootstrap";
import axios from "axios";
import "./App.css";

function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false); // Estado para el spinner de carga
  const [error, setError] = useState(null); // Estado para manejar errores

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
    setError(null); // Limpiar errores cuando se seleccione un nuevo archivo
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file) {
      setError("Por favor, selecciona un archivo PDF.");
      return;
    }

    setLoading(true); // Mostrar spinner de carga
    const formData = new FormData();
    formData.append("pdf", file);

    try {
      const response = await axios.post(
        "https://backendconverter.onrender.com/upload",
        formData,
        {
          headers: {
            "Content-Type": "multipart/form-data",
          },
          responseType: "blob",
        }
      );

      // Crear enlace de descarga para el archivo Excel
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", "converted.xlsx");
      document.body.appendChild(link);
      link.click();

      // Limpiar el archivo seleccionado y el error después de la descarga
      setFile(null);
      setError(null);
    } catch (error) {
      setError(
        "Hubo un error al convertir el archivo. Por favor, inténtalo de nuevo."
      );
      console.error("Error uploading the file:", error);
    } finally {
      setLoading(false); // Ocultar spinner de carga
    }
  };

  return (
    <Container className="mt-4">
      <h1>Bienvenido</h1>
      <h2>Convert PDF to Excel</h2>
      {error && <Alert variant="danger">{error}</Alert>}
      <Form onSubmit={handleSubmit}>
        <Form.Group>
          <Form.Label>Upload PDF</Form.Label>
          <Form.Control type="file" onChange={handleFileChange} />
        </Form.Group>
        <Button
          variant="primary"
          type="submit"
          className="mt-3"
          disabled={loading}
        >
          {loading ? (
            <>
              <Spinner
                as="span"
                animation="border"
                size="sm"
                role="status"
                aria-hidden="true"
              />
              {"  "} Convirtiendo...
            </>
          ) : (
            "Convertir a Excel"
          )}
        </Button>
      </Form>
    </Container>
  );
}

export default App;
