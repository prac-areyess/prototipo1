const express = require('express');
const cors = require('cors');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.json());

const DATA_FILE = path.join(__dirname, 'pedidos.json');

// Guardar pedido en JSON
app.post('/api/pedido', upload.single('archivo'), (req, res) => {
    const {
        dependencia, tipoDocumento, responsableSine, uuoo, uuooNombre,
        claseDocumento, procesoSine, encargado, numeroCliente
    } = req.body;
    const archivoNombre = req.file ? req.file.originalname : null;

    // Leer pedidos existentes
    let pedidos = [];
    if (fs.existsSync(DATA_FILE)) {
        pedidos = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    }

    // Agregar nuevo pedido
    pedidos.push({
        dependencia, tipo_documento: tipoDocumento, responsable_sine: responsableSine,
        uuoo, uuoo_nombre: uuooNombre, clase_documento: claseDocumento,
        proceso_sine: procesoSine, encargado, archivo_nombre: archivoNombre,
        numero_cliente: numeroCliente, fecha_registro: new Date()
    });

    // Guardar pedidos actualizados
    fs.writeFileSync(DATA_FILE, JSON.stringify(pedidos, null, 2));

    res.json({ success: true });
});

// Consultar pedidos
app.get('/api/pedidos', (req, res) => {
    // Leer pedidos existentes
    let pedidos = [];
    if (fs.existsSync(DATA_FILE)) {
        pedidos = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    }

    res.json(pedidos);
});

app.listen(3000, () => {
    console.log('Servidor escuchando en http://localhost:3000');
});