require('dotenv').config();
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const multer = require('multer');
const { v2: cloudinary } = require('cloudinary');
const { CloudinaryStorage } = require('multer-storage-cloudinary');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// 1. KONEKSI MONGO DB
const MONGO_URI = process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/surat_jalan_db';
mongoose.connect(MONGO_URI)
  .then(() => console.log('DB Connected successfully'))
  .catch(err => console.error('DB Connection Error:', err));

// 2. MODEL DATABASE (SCHEMA REVISI)
const itemSchema = new mongoose.Schema({
  nama_barang: String,
  spesifikasi: String,
  keterangan: String,
  foto: String,
  qty: Number,
  satuan: String,
  harga: Number
});

const suratJalanSchema = new mongoose.Schema({
  _id: { type: String }, // REVISI: Mengizinkan Custom String ID agar sync dengan LocalStorage
  divisi: { type: String, required: true },
  tipe: { type: String, required: true },
  tanggal: { type: String, required: true },
  no_surat: { type: String, required: true },
  customer: { type: String, required: true },
  items: [itemSchema]
}, { timestamps: true, _id: false }); // Mematikan auto-generate ObjectId bawaan jika custom ID dipakai

const SuratJalan = mongoose.model('SuratJalan', suratJalanSchema);

// 3. KONFIGURASI UPLOAD CLOUDINARY
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET
});

const storage = new CloudinaryStorage({
  cloudinary: cloudinary,
  params: {
    folder: 'surat_jalan_photos',
    allowed_formats: ['jpg', 'png', 'jpeg']
  }
});
const upload = multer({ storage: storage });

// --- ENDPOINT API ---

app.post('/api/upload', upload.single('foto'), (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'File gagal diunggah' });
  res.json({ url: req.file.path });
});

app.get('/api/surat-jalan', async (req, res) => {
  try {
    const { divisi } = req.query;
    const filter = divisi ? { divisi } : {};
    const list = await SuratJalan.find(filter).sort({ createdAt: -1 });
    res.json(list);
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

app.post('/api/surat-jalan', async (req, res) => {
  try {
    if (!req.body._id) {
      req.body._id = 'sj_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7);
    }
    const baru = new SuratJalan(req.body);
    await baru.save();
    res.status(201).json({ message: 'Surat Jalan berhasil disimpan', data: baru });
  } catch (err) {
    res.status(400).json({ message: err.message });
  }
});

app.put('/api/surat-jalan/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const updated = await SuratJalan.findByIdAndUpdate(id, req.body, { new: true });
    res.json({ message: 'Surat Jalan berhasil diperbarui', data: updated });
  } catch (err) {
    res.status(400).json({ message: err.message });
  }
});

app.delete('/api/surat-jalan/:id', async (req, res) => {
  try {
    const { id } = req.params;
    await SuratJalan.findByIdAndDelete(id);
    res.json({ message: 'Data berhasil dihapus!' });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));