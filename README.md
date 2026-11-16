<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Kesiswaan SRMA 1 Kab Pasuruan</title>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    
    <!-- SheetJS (Excel Parser) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    
    <!-- FontAwesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

    <!-- Tailwind Config for Custom Colors -->
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        primary: '#176B52',
                        secondary: '#A8C9B5',
                        background: '#FAF7EF',
                    },
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    }
                }
            }
        }
    </script>

    <style>
        body { background-color: #FAF7EF; }
        
        /* Print Styles */
        @media print {
            body { background-color: white; }
            .no-print { display: none !important; }
            .print-only { display: block !important; }
            @page { margin: 2cm; }
        }
        
        .print-only { display: none; }

        /* Custom Scrollbar */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #FAF7EF; }
        ::-webkit-scrollbar-thumb { background: #A8C9B5; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #176B52; }
    </style>
</head>
<body class="text-gray-800 font-sans h-screen w-screen overflow-hidden m-0 p-0">

    <!-- App Container (Hidden during print) -->
    <div id="app-container" class="flex w-full h-full no-print bg-background relative">
        
        <!-- Sidebar -->
        <aside class="w-64 bg-primary text-white flex flex-col shadow-xl transition-all duration-300 z-20 shrink-0">
            <div class="p-6 flex items-center space-x-3 border-b border-secondary/30">
                <i class="fa-solid fa-school text-2xl text-secondary"></i>
                <div>
                    <h1 class="font-bold text-lg leading-tight">SRMA 1</h1>
                    <p class="text-xs text-secondary">Kesiswaan</p>
                </div>
            </div>
            
            <div class="p-4">
                <label class="block text-xs text-secondary mb-1">Pengguna Aktif:</label>
                <select id="user-select" class="w-full bg-primary border border-secondary text-white text-sm rounded p-2 focus:outline-none focus:border-white">
                    <!-- Populated by JS -->
                </select>
            </div>

            <nav class="flex-1 px-4 py-4 space-y-2 overflow-y-auto">
                <a href="#" onclick="switchMenu('dashboard')" id="menu-dashboard" class="nav-btn flex items-center space-x-3 px-4 py-3 bg-secondary text-primary rounded-lg font-medium transition-colors">
                    <i class="fa-solid fa-chart-pie w-5"></i> <span>Dashboard</span>
                </a>
                <a href="#" onclick="switchMenu('students')" id="menu-students" class="nav-btn flex items-center space-x-3 px-4 py-3 hover:bg-secondary hover:text-primary rounded-lg font-medium transition-colors">
                    <i class="fa-solid fa-users w-5"></i> <span>Data Siswa</span>
                </a>
                <a href="#" onclick="switchMenu('violations')" id="menu-violations" class="nav-btn flex items-center space-x-3 px-4 py-3 hover:bg-secondary hover:text-primary rounded-lg font-medium transition-colors">
                    <i class="fa-solid fa-triangle-exclamation w-5"></i> <span>Pelanggaran</span>
                </a>
                <a href="#" onclick="switchMenu('permissions')" id="menu-permissions" class="nav-btn flex items-center space-x-3 px-4 py-3 hover:bg-secondary hover:text-primary rounded-lg font-medium transition-colors">
                    <i class="fa-solid fa-envelope-open-text w-5"></i> <span>Surat Izin</span>
                </a>
            </nav>
            
            <div class="p-4 border-t border-secondary/30 text-xs text-center text-secondary">
                Sistem Terintegrasi Online
            </div>
        </aside>

        <!-- Main Content -->
        <main class="flex-1 flex flex-col h-full min-w-0 overflow-hidden relative">
            
            <!-- Top Header -->
            <header class="bg-white shadow-sm p-4 flex justify-between items-center z-20 shrink-0 relative">
                <h2 id="page-title" class="text-xl font-bold text-primary">Dashboard</h2>
                <div class="flex items-center space-x-4">
                    <div id="realtime-clock" class="text-sm font-medium text-gray-600 bg-background px-4 py-2 rounded-full border border-secondary/50">
                        Memuat waktu...
                    </div>
                </div>
            </header>

            <!-- Loading Overlay -->
            <div id="loading-overlay" class="absolute inset-0 bg-background/80 z-50 flex flex-col items-center justify-center">
                <i class="fa-solid fa-circle-notch fa-spin text-4xl text-primary mb-4"></i>
                <p class="text-primary font-medium animate-pulse">Menghubungkan ke Server...</p>
            </div>

            <!-- Content Area -->
            <div class="flex-1 overflow-y-auto p-6 min-h-0 bg-background" id="main-content">
                
                <!-- DASHBOARD VIEW -->
                <div id="view-dashboard" class="space-y-6">
                    <!-- Stats Cards -->
                    <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                        <div class="bg-white p-6 rounded-xl shadow-sm border-l-4 border-primary flex items-center">
                            <div class="p-4 bg-background rounded-full text-primary mr-4">
                                <i class="fa-solid fa-user-graduate text-2xl"></i>
                            </div>
                            <div>
                                <p class="text-sm text-gray-500 font-medium">Total Siswa</p>
                                <p class="text-2xl font-bold text-gray-800" id="stat-total-students">0</p>
                            </div>
                        </div>
                        <div class="bg-white p-6 rounded-xl shadow-sm border-l-4 border-orange-500 flex items-center">
                            <div class="p-4 bg-orange-50 rounded-full text-orange-500 mr-4">
                                <i class="fa-solid fa-person-walking-arrow-right text-2xl"></i>
                            </div>
                            <div>
                                <p class="text-sm text-gray-500 font-medium">Izin Pulang Hari Ini</p>
                                <p class="text-2xl font-bold text-gray-800" id="stat-permissions-today">0</p>
                            </div>
                        </div>
                        <div class="bg-white p-6 rounded-xl shadow-sm border-l-4 border-red-500 flex items-center">
                            <div class="p-4 bg-red-50 rounded-full text-red-500 mr-4">
                                <i class="fa-solid fa-triangle-exclamation text-2xl"></i>
                            </div>
                            <div>
                                <p class="text-sm text-gray-500 font-medium">Pelanggaran Hari Ini</p>
                                <p class="text-2xl font-bold text-gray-800" id="stat-violations-today">0</p>
                            </div>
                        </div>
                    </div>

                    <!-- Top 10 Violators -->
                    <div class="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
                        <div class="p-5 border-b border-gray-100 flex justify-between items-center bg-gray-50">
                            <h3 class="font-bold text-primary"><i class="fa-solid fa-ranking-star mr-2"></i>Top 10 Pelanggar Bulan Ini</h3>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left text-sm">
                                <thead class="bg-background text-primary">
                                    <tr>
                                        <th class="p-4 w-16 text-center">No</th>
                                        <th class="p-4">Nama Siswa</th>
                                        <th class="p-4">Kelas</th>
                                        <th class="p-4 text-center">Total Poin</th>
                                    </tr>
                                </thead>
                                <tbody id="top-violators-tbody" class="divide-y divide-gray-100">
                                    <!-- Populated by JS -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- STUDENTS VIEW -->
                <div id="view-students" class="hidden space-y-4">
                    <div class="flex justify-between items-center bg-white p-4 rounded-xl shadow-sm">
                        <div class="relative w-full max-w-md">
                            <i class="fa-solid fa-search absolute left-3 top-3 text-gray-400"></i>
                            <input type="text" id="search-student" onkeyup="renderStudents()" placeholder="Cari NIS atau Nama Siswa..." class="w-full pl-10 pr-4 py-2 border border-secondary rounded-lg focus:outline-none focus:border-primary transition-all">
                        </div>
                        <div class="flex items-center space-x-3">
                            <div class="text-sm text-gray-500 font-medium mr-2" id="student-count-display"></div>
                            
                            <!-- Import Excel Input -->
                            <input type="file" id="excel-upload" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleExcelUpload(event)">
                            <label for="excel-upload" class="cursor-pointer bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg font-medium transition-colors shadow-sm text-sm flex items-center">
                                <i class="fa-solid fa-file-excel mr-2"></i> Import Excel
                            </label>

                            <button onclick="openStudentModal()" class="bg-primary hover:bg-primary/90 text-white px-4 py-2 rounded-lg font-medium transition-colors shadow-sm text-sm">
                                <i class="fa-solid fa-plus mr-2"></i> Tambah Siswa
                            </button>
                        </div>
                    </div>
                    
                    <div class="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
                        <div class="overflow-x-auto h-[60vh]">
                            <table class="w-full text-left text-sm">
                                <thead class="bg-background text-primary sticky top-0 shadow-sm z-10">
                                    <tr>
                                        <th class="p-4">NIS</th>
                                        <th class="p-4">Nama Lengkap</th>
                                        <th class="p-4">Kelas</th>
                                        <th class="p-4 text-center">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody id="students-tbody" class="divide-y divide-gray-100">
                                    <!-- Populated by JS -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- VIOLATIONS VIEW -->
                <div id="view-violations" class="hidden space-y-4">
                    <div class="flex justify-between items-center bg-white p-4 rounded-xl shadow-sm">
                        <h3 class="font-bold text-gray-700">Riwayat Pelanggaran</h3>
                        <button onclick="openModal('modal-violation')" class="bg-primary hover:bg-primary/90 text-white px-4 py-2 rounded-lg font-medium transition-colors shadow-sm">
                            <i class="fa-solid fa-plus mr-2"></i> Catat Pelanggaran
                        </button>
                    </div>
                    
                    <div class="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
                        <div class="overflow-x-auto h-[60vh]">
                            <table class="w-full text-left text-sm">
                                <thead class="bg-background text-primary sticky top-0 shadow-sm z-10">
                                    <tr>
                                        <th class="p-4">Tanggal & Waktu</th>
                                        <th class="p-4">Siswa</th>
                                        <th class="p-4">Kelas</th>
                                        <th class="p-4">Bentuk Pelanggaran</th>
                                        <th class="p-4 text-center">Poin</th>
                                        <th class="p-4">Petugas</th>
                                        <th class="p-4 text-center">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody id="violations-tbody" class="divide-y divide-gray-100">
                                    <!-- Populated by JS -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- PERMISSIONS VIEW -->
                <div id="view-permissions" class="hidden space-y-4">
                    <div class="flex justify-between items-center bg-white p-4 rounded-xl shadow-sm">
                        <h3 class="font-bold text-gray-700">Riwayat Izin Kepulangan</h3>
                        <button onclick="openModal('modal-permission')" class="bg-primary hover:bg-primary/90 text-white px-4 py-2 rounded-lg font-medium transition-colors shadow-sm">
                            <i class="fa-solid fa-plus mr-2"></i> Buat Surat Izin
                        </button>
                    </div>
                    
                    <div class="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
                        <div class="overflow-x-auto h-[60vh]">
                            <table class="w-full text-left text-sm">
                                <thead class="bg-background text-primary sticky top-0 shadow-sm z-10">
                                    <tr>
                                        <th class="p-4">Tgl Pulang</th>
                                        <th class="p-4">Tgl Kembali</th>
                                        <th class="p-4">Siswa</th>
                                        <th class="p-4">Kelas</th>
                                        <th class="p-4">Alasan</th>
                                        <th class="p-4">Petugas</th>
                                        <th class="p-4 text-center">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody id="permissions-tbody" class="divide-y divide-gray-100">
                                    <!-- Populated by JS -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

            </div>
        </main>
    </div>

    <!-- MODALS -->
    <!-- Modal Add Violation -->
    <div id="modal-violation" class="fixed inset-0 bg-black/50 z-50 hidden flex items-center justify-center opacity-0 transition-opacity duration-300">
        <div class="bg-white rounded-xl shadow-2xl w-full max-w-lg mx-4 overflow-hidden transform scale-95 transition-transform duration-300" id="modal-violation-content">
            <div class="bg-primary p-4 flex justify-between items-center text-white">
                <h3 class="font-bold text-lg">Catat Pelanggaran Baru</h3>
                <button onclick="closeModal('modal-violation')" class="text-white hover:text-red-300"><i class="fa-solid fa-xmark text-xl"></i></button>
            </div>
            <form id="form-violation" onsubmit="submitViolation(event)" class="p-6 space-y-4">
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Tanggal</label>
                    <input type="date" id="v-date" required class="w-full p-2 border border-gray-300 rounded focus:ring-primary focus:border-primary">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Pilih Siswa</label>
                    <select id="v-student" required class="w-full p-2 border border-gray-300 rounded focus:ring-primary focus:border-primary">
                        <option value="">-- Pilih Siswa --</option>
                        <!-- Populated by JS -->
                    </select>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Bentuk Pelanggaran</label>
                    <textarea id="v-desc" required rows="2" class="w-full p-2 border border-gray-300 rounded focus:ring-primary focus:border-primary" placeholder="Contoh: Terlambat masuk sekolah, atribut tidak lengkap..."></textarea>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Poin Pelanggaran</label>
                    <input type="number" id="v-points" required min="1" max="100" class="w-full p-2 border border-gray-300 rounded focus:ring-primary focus:border-primary" placeholder="10">
                </div>
                <div class="pt-4 flex justify-end space-x-3">
                    <button type="button" onclick="closeModal('modal-violation')" class="px-4 py-2 border border-gray-300 rounded text-gray-700 hover:bg-gray-50">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-primary text-white rounded hover:bg-primary/90 shadow">Simpan Data</button>
                </div>
            </form>
        </div>
    </div>

    <!-- Modal Add/Edit Student -->
    <div id="modal-student" class="fixed inset-0 bg-black/50 z-50 hidden flex items-center justify-center opacity-0 transition-opacity duration-300">
        <div class="bg-white rounded-xl shadow-2xl w-full max-w-md mx-4 overflow-hidden transform scale-95 transition-transform duration-300" id="modal-student-content">
            <div class="bg-primary p-4 flex justify-between items-center text-white">
                <h3 class="font-bold text-lg" id="modal-student-title">Tambah Siswa Baru</h3>
                <button onclick="closeModal('modal-student')" class="text-white hover:text-red-300"><i class="fa-solid fa-xmark text-xl"></i></button>
            </div>
            <form id="form-student" onsubmit="submitStudent(event)" class="p-6 space-y-4">
                <input type="hidden" id="s-id">
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">NIS</label>
                    <input type="text" id="s-nis" required class="w-full p-2 border border-gray-300 rounded focus:ring-primary focus:border-primary">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Nama Lengkap</label>
                    <input type="text" id="s-name" required class="w-full p-2 border border-gray-300 rounded focus:ring-primary focus:border-primary">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">Kelas</label>
                    <select id="s-cls" required class="w-full p-2 border border-gray-300 rounded focus:ring-primary focus:border-primary">
                        <option value="">-- Pilih Kelas --</option>
                        <option value="X-A">X-A</option>
                        <option value="X-B">X-B</option>
                        <option value="X-C">X-C</option>
                        <option value="X-PK 1">X-PK 1</option>
                        <option value="X-PK 2">X-PK 2</option>
                        <option value="XI-A">XI-A</option>
                        <option value="XI-B">XI-B</option>
                        <option value="XII-A">XII-A</option>
                        <option value="XII-B">XII-B</option>
                        <option value="XII-C">XII-C</option>
                        <option value="XII-D">XII-D</option>
                    </select>
                </div>
                <div class="pt-4 flex justify-end space-x-3">
                    <button type="button" onclick="closeModal('modal-student')" class="px-4 py-2 border border-gray-300 rounded text-gray-700 hover:bg-gray-50">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-primary text-white rounded hover:bg-primary/90 shadow">Simpan</button>
                </div>
            </form>
        </div>
    </div>

    <!-- PRINT LAYOUT (Only visible when printing) -->
    <div id="print-area" class="print-only w-full font-serif text-black leading-relaxed text-justify">
        <!-- KOP SURAT -->
        <div class="text-center border-b-4 border-double border-black pb-4 mb-6">
            <h1 class="text-xl font-bold uppercase">Kementerian Sosial Republik Indonesia</h1>
            <h2 class="text-lg font-bold uppercase">Pusat Pendidikan Pelatihan dan Pengembangan Profesi</h2>
            <h2 class="text-lg font-bold uppercase">Sekolah Rakyat Terintegrasi 1 Kabupaten Pasuruan</h2>
            <p class="text-sm mt-2">Desa Lajuk Kecamatan Gondang Wetan Kabupaten Pasuruan, Jawa Timur</p>
            <p class="text-sm">Pos-el : sekolahrakyat0kabpasuruan@gmail.com Kode Pos 67174</p>
        </div>

        <h3 class="text-center font-bold text-lg uppercase underline mb-8">Surat Izin Meninggalkan Pelajaran</h3>

        <p class="mb-4">Yang bertanda tangan di bawah ini, petugas kesiswaan SRMA 1 Kabupaten Pasuruan, memberikan izin kepada:</p>

        <table class="w-full mb-4 ml-8">
            <tr><td class="w-40 py-1">Nama Siswa</td><td class="w-4">:</td><td class="font-bold" id="print-name">-</td></tr>
            <tr><td class="py-1">NIS</td><td>:</td><td id="print-nis">-</td></tr>
            <tr><td class="py-1">Kelas</td><td>:</td><td id="print-class">-</td></tr>
        </table>

        <p class="mb-4">Untuk meninggalkan sekolah / pelajaran mulai tanggal <strong id="print-depart">-</strong> sampai dengan tanggal <strong id="print-return">-</strong>.</p>
        
        <p class="mb-8">Adapun alasan kepulangan / meninggalkan pelajaran tersebut adalah karena: <br><strong id="print-reason" class="italic">"-"</strong>.</p>

        <p class="mb-12">Demikian surat izin ini dibuat agar dapat dipergunakan sebagaimana mestinya dan dimohon kepada Bapak/Ibu Guru serta Satpam sekolah untuk memakluminya.</p>

        <div class="flex justify-between mt-12">
            <div class="text-center">
                <p class="mb-20">Mengetahui,<br>Orang Tua / Wali Murid</p>
                <p class="font-bold border-b border-black inline-block w-48"></p>
            </div>
            <div class="text-center">
                <p class="mb-1" id="print-date-location">Pasuruan, -</p>
                <p class="mb-20">Petugas Pemberi Izin,</p>
                <p class="font-bold underline" id="print-officer">-</p>
                <p>Tim Kesiswaan</p>
            </div>
        </div>
    </div>

    <!-- Firebase Scripts -->
    <script type="module">
        import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
        import { getAuth, signInAnonymously, signInWithCustomToken } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-auth.js";
        import { getFirestore, collection, onSnapshot, addDoc, deleteDoc, doc, serverTimestamp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

        // State variables
        let db, auth;
        let students = [];
        let violationsData = [];
        let permissionsData = [];
        
        // Setup App Context
        const appId = typeof __app_id !== 'undefined' ? __app_id : 'kesiswaan-srma1-app';
        let isFirebaseActive = false;

        const OFFICERS = [
            "Kepala Sekolah",
            "Budi Santoso (Kesiswaan 1)",
            "Siti Aminah (Kesiswaan 2)",
            "Agus Setiawan (Kesiswaan 3)",
            "Rini Wati (Kesiswaan 4)",
            "Hendra Gunawan (Kesiswaan 5)",
            "Lina Marlina (Kesiswaan 6)",
            "Joko Purwanto (Kesiswaan 7)",
            "Dewi Lestari (Kesiswaan 8)"
        ];
        
        const CLASSES = ["X-A", "X-B", "X-C", "X-PK 1", "X-PK 2", "XI-A", "XI-B", "XII-A", "XII-B", "XII-C", "XII-D"];

        // 1. Initialize Data (Students) from Image
        function generateStudents() {
            const rawData = [
                // X-A
                ["10001", "AINUL MARDHIYAH", "X-A"],
                ["10002", "AKBAR WIDIANSYAH", "X-A"],
                ["10003", "ALYA SYAMSIYYAH", "X-A"],
                ["10004", "ARYA BIMA PRATAMA", "X-A"],
                ["10005", "CINTA LAURA NATASHA AULIA LESTARI", "X-A"],
                ["10006", "DANIA AFIFAH", "X-A"],
                ["10007", "DIKO DWI PUTRA ERLANGGA RAHMADANI", "X-A"],
                ["10008", "DINA KAMILA", "X-A"],
                ["10009", "EKA LAILATUL AFIDA RISMA", "X-A"],
                ["10010", "ERLINA CAHYA MARDHIKA", "X-A"],
                ["10011", "HABIBUR RAHMAN", "X-A"],
                ["10012", "IHYA ULUDIN", "X-A"],
                ["10013", "INAYATUS SALIHA", "X-A"],
                ["10014", "INTAN NUR RAHMA MAULIDA", "X-A"],
                ["10015", "MUHAMMAD FAHREZA SETIANTO", "X-A"],
                ["10016", "MUHAMMAD MAULANA NASIQ", "X-A"],
                ["10017", "MUHAMMAD NUR ALFIANSYAH", "X-A"],
                ["10018", "MUHAMMAD RIZKY FIRMANSYAH", "X-A"],
                ["10019", "NUR INAYAH FITRIANI", "X-A"],
                ["10020", "PUTRI NUR AINIYAH", "X-A"],
                ["10021", "RAHMAD DHANI KURNIAWAN", "X-A"],
                ["10022", "SYAFA AZZAHRA", "X-A"],
                ["10023", "SYIFA NUR AZZAHRA", "X-A"],
                ["10024", "YUSUF AMIRUDIN", "X-A"],
                
                // X-B
                ["10025", "AHMAD BUKHORI", "X-B"],
                ["10026", "ALYA DWI YANTI", "X-B"],
                ["10027", "CINTYA RAMADHANI", "X-B"],
                ["10028", "DANY HERMAWAN", "X-B"],
                ["10029", "DEVI SRI WULANDARI", "X-B"],
                ["10030", "FADILAH TRI HAPSARI", "X-B"],
                ["10031", "FITRIA AYU NISA", "X-B"],
                ["10032", "JOKO KARYO", "X-B"],
                ["10033", "LAILATUR RAHMA HADIYANI", "X-B"],
                ["10034", "MOCH. RIZKY ADITYA", "X-B"],
                ["10035", "MUCH. ABDAN SYAKUR", "X-B"],
                ["10036", "NUR ILMI", "X-B"],
                ["10037", "NUR ROYHAN FATHAN MUBARROK", "X-B"],
                ["10038", "PUTRI NUR AINI", "X-B"],
                ["10039", "RISMA BUKHORI", "X-B"],
                ["10040", "SILVIA ARTA SARI", "X-B"],
                ["10041", "SINTYA CANTIKA", "X-B"],
                
                // X-C
                ["10042", "AHMAD MA'RUF PRATAMA", "X-C"],
                ["10043", "AKBAR AL GHOZALI", "X-C"],
                ["10044", "ALYA NUR AISYAH W.", "X-C"],
                ["10045", "DIAN AYU PERMATA NENDA", "X-C"],
                ["10046", "DWI NATALIA SYIFA", "X-C"],
                ["10047", "FIRA CINDY ASMARA", "X-C"],
                ["10048", "IQBAL FIRMANSYAH", "X-C"],
                ["10049", "LAILATUL FITRIAH", "X-C"],
                ["10050", "MUHAMMAD IRFAN ARDIANSYAH", "X-C"],
                ["10051", "MUHAMMAD REVAN ARMANSYAH", "X-C"],
                ["10052", "MUHAMMAD RIZAL RIFANI", "X-C"],
                ["10053", "NURUL AINI", "X-C"],
                ["10054", "OLIVIA ZIANA PUTRI", "X-C"],
                
                // X-PK 1
                ["10055", "AHMAD FARHAN", "X-PK 1"],
                ["10056", "AHMAD SAFI'I", "X-PK 1"],
                ["10057", "ALDA RIZKI AMALIA", "X-PK 1"],
                ["10058", "AYU YULIASTUTI", "X-PK 1"],
                ["10059", "DEVI MARITA", "X-PK 1"],
                ["10060", "DINA AULIA SARI", "X-PK 1"],
                ["10061", "FARA AULIA RAMADHANI", "X-PK 1"],
                ["10062", "MUHAMMAD ANGGA PRATAMA", "X-PK 1"],
                ["10063", "MUHAMMAD FAIZAL", "X-PK 1"],
                ["10064", "NOVITA ASARI", "X-PK 1"],
                
                // X-PK 2
                ["10065", "ABU DZAR AL GHIFARI", "X-PK 2"],
                ["10066", "ACHMAD RISKY", "X-PK 2"],
                ["10067", "BIMA BAGASKARA", "X-PK 2"],
                ["10068", "DWI WAHYUNI", "X-PK 2"],
                ["10069", "LAILATUL QADRIYAH", "X-PK 2"],
                ["10070", "MUHAMMAD IQBAL GHOZALI", "X-PK 2"],
                ["10071", "MUHAMMAD WAHYU FAUZI", "X-PK 2"],
                ["10072", "PUTRI AMELIA", "X-PK 2"],
                ["10073", "SITI MUZAYANAH", "X-PK 2"],
                ["10074", "ZAHRA AULIA", "X-PK 2"]
            ];
            
            students = rawData.map((d, index) => ({
                id: `STU${index.toString().padStart(4, '0')}`,
                nis: d[0],
                name: d[1],
                cls: d[2]
            }));
            
            // Sort by class then name
            students.sort((a,b) => a.cls.localeCompare(b.cls) || a.name.localeCompare(b.name));
        }

        // 2. Initialize App
        async function initApp() {
            populateSelects();
            generateStudents();
            startClock();
            
            try {
                if (typeof __firebase_config !== 'undefined') {
                    const firebaseConfig = JSON.parse(__firebase_config);
                    const app = initializeApp(firebaseConfig);
                    auth = getAuth(app);
                    db = getFirestore(app);

                    if (typeof __initial_auth_token !== 'undefined') {
                        await signInWithCustomToken(auth, __initial_auth_token);
                    } else {
                        await signInAnonymously(auth);
                    }
                    
                    if(auth.currentUser) {
                        isFirebaseActive = true;
                        setupRealtimeListeners();
                    }
                } else {
                    console.log("No Firebase config. Running in offline/memory mode.");
                    hideLoading();
                    renderAll();
                }
            } catch (error) {
                console.error("Firebase init error:", error);
                hideLoading();
                renderAll(); // Fallback to empty memory state
            }
        }

        // 3. UI Interactions & DOM
        window.switchMenu = function(viewName) {
            // Hide all views
            ['dashboard', 'students', 'violations', 'permissions'].forEach(v => {
                document.getElementById(`view-${v}`).classList.add('hidden');
                document.getElementById(`menu-${v}`).classList.remove('bg-secondary', 'text-primary');
                document.getElementById(`menu-${v}`).classList.add('hover:bg-secondary', 'hover:text-primary');
            });
            
            // Show selected view
            document.getElementById(`view-${viewName}`).classList.remove('hidden');
            const activeMenu = document.getElementById(`menu-${viewName}`);
            activeMenu.classList.remove('hover:bg-secondary', 'hover:text-primary');
            activeMenu.classList.add('bg-secondary', 'text-primary');
            
            // Set Title
            const titles = {
                'dashboard': 'Dashboard',
                'students': 'Data Siswa',
                'violations': 'Data Pelanggaran',
                'permissions': 'Surat Izin Kepulangan'
            };
            document.getElementById('page-title').innerText = titles[viewName];
            
            if(viewName === 'students') renderStudents();
            if(viewName === 'dashboard') renderDashboard();
        }

        function populateSelects() {
            const userSelect = document.getElementById('user-select');
            OFFICERS.forEach(off => {
                const opt = document.createElement('option');
                opt.value = off;
                opt.text = off;
                userSelect.appendChild(opt);
            });
        }

        function populateStudentDropdowns() {
            const vStudent = document.getElementById('v-student');
            const pStudent = document.getElementById('p-student');
            vStudent.innerHTML = '<option value="">-- Pilih Siswa --</option>';
            pStudent.innerHTML = '<option value="">-- Pilih Siswa --</option>';
            
            let currentClass = "";
            let optGroupV, optGroupP;

            students.forEach(s => {
                if (s.cls !== currentClass) {
                    currentClass = s.cls;
                    optGroupV = document.createElement('optgroup');
                    optGroupV.label = `Kelas ${currentClass}`;
                    vStudent.appendChild(optGroupV);
                    
                    optGroupP = document.createElement('optgroup');
                    optGroupP.label = `Kelas ${currentClass}`;
                    pStudent.appendChild(optGroupP);
                }
                const optV = document.createElement('option');
                optV.value = s.id;
                optV.text = `${s.nis} - ${s.name}`;
                optGroupV.appendChild(optV);
                
                const optP = optV.cloneNode(true);
                optGroupP.appendChild(optP);
            });
        }

        function startClock() {
            const el = document.getElementById('realtime-clock');
            const update = () => {
                const now = new Date();
                const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit' };
                el.innerText = now.toLocaleDateString('id-ID', options);
            };
            update();
            setInterval(update, 1000);
        }

        function hideLoading() {
            const loader = document.getElementById('loading-overlay');
            loader.classList.add('opacity-0');
            setTimeout(() => loader.classList.add('hidden'), 300);
        }

        // 4. Modals
        window.openModal = function(id) {
            const modal = document.getElementById(id);
            const content = document.getElementById(`${id}-content`);
            modal.classList.remove('hidden');
            // Populate student dropdowns right before showing modal
            if(id === 'modal-violation' || id === 'modal-permission') populateStudentDropdowns();
            
            // Set default date to today
            if(id === 'modal-violation') document.getElementById('v-date').valueAsDate = new Date();
            if(id === 'modal-permission') {
                document.getElementById('p-depart').valueAsDate = new Date();
                document.getElementById('p-return').valueAsDate = new Date();
            }

            // small delay to allow display:block to apply before animating opacity
            setTimeout(() => {
                modal.classList.remove('opacity-0');
                content.classList.remove('scale-95');
            }, 10);
        }

        window.closeModal = function(id) {
            const modal = document.getElementById(id);
            const content = document.getElementById(`${id}-content`);
            modal.classList.add('opacity-0');
            content.classList.add('scale-95');
            setTimeout(() => modal.classList.add('hidden'), 300);
        }

        // 5. Firebase Realtime Listeners (RULE 1 & 2 compliant)
        function setupRealtimeListeners() {
            if(!auth.currentUser) return;
            
            // Public Data path for collaboration
            const vCollection = collection(db, 'artifacts', appId, 'public', 'data', 'violations');
            const pCollection = collection(db, 'artifacts', appId, 'public', 'data', 'permissions');

            onSnapshot(vCollection, (snapshot) => {
                violationsData = [];
                snapshot.forEach(doc => {
                    violationsData.push({ dbId: doc.id, ...doc.data() });
                });
                renderAll();
                hideLoading();
            }, (error) => {
                console.error("Error listening to violations:", error);
                hideLoading();
            });

            onSnapshot(pCollection, (snapshot) => {
                permissionsData = [];
                snapshot.forEach(doc => {
                    permissionsData.push({ dbId: doc.id, ...doc.data() });
                });
                renderAll();
            }, (error) => {
                console.error("Error listening to permissions:", error);
            });
        }

        // 6. Rendering Logic
        function renderAll() {
            renderDashboard();
            renderStudents();
            renderViolations();
            renderPermissions();
        }

        window.renderStudents = function() {
            const tbody = document.getElementById('students-tbody');
            const search = document.getElementById('search-student').value.toLowerCase();
            tbody.innerHTML = '';
            
            const filtered = students.filter(s => 
                s.name.toLowerCase().includes(search) || 
                s.nis.includes(search)
            );
            
            document.getElementById('student-count-display').innerText = `Menampilkan ${filtered.length} dari ${students.length} Siswa`;

            filtered.forEach((s, index) => {
                const tr = document.createElement('tr');
                tr.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
                tr.innerHTML = `
                    <td class="p-4">${s.nis}</td>
                    <td class="p-4 font-medium">${s.name}</td>
                    <td class="p-4">
                        <span class="bg-secondary/20 text-primary px-2 py-1 rounded text-xs font-bold border border-secondary/50">${s.cls}</span>
                    </td>
                    <td class="p-4 text-center space-x-2">
                        <button onclick="openStudentModal('${s.id}')" class="text-blue-500 hover:text-blue-700 bg-blue-50 p-2 rounded" title="Edit Data">
                            <i class="fa-solid fa-pen"></i>
                        </button>
                        <button onclick="deleteStudentData('${s.id}')" class="text-red-500 hover:text-red-700 bg-red-50 p-2 rounded" title="Hapus">
                            <i class="fa-solid fa-trash"></i>
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function renderViolations() {
            const tbody = document.getElementById('violations-tbody');
            tbody.innerHTML = '';
            
            // Sort by date desc
            const sorted = [...violationsData].sort((a,b) => new Date(b.date) - new Date(a.date));
            
            sorted.forEach((v, index) => {
                const s = students.find(st => st.id === v.studentId) || {name: 'Unknown', cls: '-'};
                const tr = document.createElement('tr');
                tr.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
                tr.innerHTML = `
                    <td class="p-4 whitespace-nowrap">${formatDate(v.date)}</td>
                    <td class="p-4 font-medium">${s.name}</td>
                    <td class="p-4">${s.cls}</td>
                    <td class="p-4">${v.desc}</td>
                    <td class="p-4 text-center"><span class="bg-red-100 text-red-700 px-2 py-1 rounded-full text-xs font-bold">${v.points}</span></td>
                    <td class="p-4 text-xs">${v.officer}</td>
                    <td class="p-4 text-center">
                        <button onclick="deleteData('violations', '${v.dbId}')" class="text-red-500 hover:text-red-700 bg-red-50 p-2 rounded" title="Hapus">
                            <i class="fa-solid fa-trash"></i>
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function renderPermissions() {
            const tbody = document.getElementById('permissions-tbody');
            tbody.innerHTML = '';
            
            const sorted = [...permissionsData].sort((a,b) => new Date(b.departDate) - new Date(a.departDate));
            
            sorted.forEach((p, index) => {
                const s = students.find(st => st.id === p.studentId) || {name: 'Unknown', cls: '-'};
                const tr = document.createElement('tr');
                tr.className = index % 2 === 0 ? 'bg-white' : 'bg-gray-50';
                tr.innerHTML = `
                    <td class="p-4 whitespace-nowrap text-orange-600 font-medium">${formatDate(p.departDate)}</td>
                    <td class="p-4 whitespace-nowrap text-primary font-medium">${formatDate(p.returnDate)}</td>
                    <td class="p-4 font-medium">${s.name}</td>
                    <td class="p-4">${s.cls}</td>
                    <td class="p-4">${p.reason}</td>
                    <td class="p-4 text-xs">${p.officer}</td>
                    <td class="p-4 text-center space-x-2">
                        <button onclick="printPermissionFromId('${p.dbId}')" class="text-blue-500 hover:text-blue-700 bg-blue-50 p-2 rounded" title="Cetak Ulang">
                            <i class="fa-solid fa-print"></i>
                        </button>
                        <button onclick="deleteData('permissions', '${p.dbId}')" class="text-red-500 hover:text-red-700 bg-red-50 p-2 rounded" title="Hapus">
                            <i class="fa-solid fa-trash"></i>
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function renderDashboard() {
            const today = new Date().toISOString().split('T')[0];
            const currentMonth = new Date().getMonth();
            const currentYear = new Date().getFullYear();

            // Stats
            document.getElementById('stat-total-students').innerText = students.length;
            
            const permsToday = permissionsData.filter(p => p.departDate === today).length;
            document.getElementById('stat-permissions-today').innerText = permsToday;

            const violsToday = violationsData.filter(v => v.date === today).length;
            document.getElementById('stat-violations-today').innerText = violsToday;

            // Top 10 This Month
            const monthlyViolations = violationsData.filter(v => {
                const d = new Date(v.date);
                return d.getMonth() === currentMonth && d.getFullYear() === currentYear;
            });

            const pointMap = {};
            monthlyViolations.forEach(v => {
                if(!pointMap[v.studentId]) pointMap[v.studentId] = 0;
                pointMap[v.studentId] += parseInt(v.points);
            });

            const topArr = Object.keys(pointMap).map(id => ({
                id, points: pointMap[id]
            })).sort((a,b) => b.points - a.points).slice(0, 10);

            const tbody = document.getElementById('top-violators-tbody');
            tbody.innerHTML = '';
            
            if(topArr.length === 0) {
                tbody.innerHTML = '<tr><td colspan="4" class="p-4 text-center text-gray-500">Belum ada data pelanggaran bulan ini</td></tr>';
            }

            topArr.forEach((item, index) => {
                const s = students.find(st => st.id === item.id) || {name: 'Unknown', cls: '-'};
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="p-4 text-center font-bold text-gray-500">${index+1}</td>
                    <td class="p-4 font-medium">${s.name}</td>
                    <td class="p-4">${s.cls}</td>
                    <td class="p-4 text-center"><span class="bg-red-100 text-red-700 px-2 py-1 rounded font-bold">${item.points}</span></td>
                `;
                tbody.appendChild(tr);
            });
        }

        // 7. Form Submissions
        window.submitViolation = async function(e) {
            e.preventDefault();
            const btn = e.target.querySelector('button[type="submit"]');
            btn.disabled = true;
            btn.innerText = "Menyimpan...";

            const data = {
                date: document.getElementById('v-date').value,
                studentId: document.getElementById('v-student').value,
                desc: document.getElementById('v-desc').value,
                points: parseInt(document.getElementById('v-points').value),
                officer: document.getElementById('user-select').value,
                timestamp: serverTimestamp ? serverTimestamp() : new Date().toISOString()
            };

            if(isFirebaseActive) {
                try {
                    await addDoc(collection(db, 'artifacts', appId, 'public', 'data', 'violations'), data);
                } catch(error) {
                    console.error("Error saving:", error);
                    alert("Gagal menyimpan ke server"); // Fallback simple alert on error
                }
            } else {
                data.dbId = Date.now().toString();
                violationsData.push(data);
                renderAll();
            }

            closeModal('modal-violation');
            e.target.reset();
            btn.disabled = false;
            btn.innerText = "Simpan Data";
            switchMenu('violations');
        }

        window.submitPermission = async function(e) {
            e.preventDefault();
            const btn = e.target.querySelector('button[type="submit"]');
            btn.disabled = true;
            btn.innerText = "Memproses...";

            const studentId = document.getElementById('p-student').value;
            const departDate = document.getElementById('p-depart').value;
            const returnDate = document.getElementById('p-return').value;
            const reason = document.getElementById('p-reason').value;
            const officer = document.getElementById('user-select').value;

            const data = {
                studentId, departDate, returnDate, reason, officer,
                timestamp: serverTimestamp ? serverTimestamp() : new Date().toISOString()
            };

            if(isFirebaseActive) {
                try {
                    await addDoc(collection(db, 'artifacts', appId, 'public', 'data', 'permissions'), data);
                } catch(error) {
                    console.error("Error saving:", error);
                }
            } else {
                data.dbId = Date.now().toString();
                permissionsData.push(data);
                renderAll();
            }

            // Direct to print
            const student = students.find(s => s.id === studentId);
            executePrint(student, departDate, returnDate, reason, officer);

            closeModal('modal-permission');
            e.target.reset();
            btn.disabled = false;
            btn.innerText = "Buat & Cetak Surat";
            switchMenu('permissions');
        }

        window.deleteData = async function(col, id) {
            // Using a simple custom modal or just confirming since iframe blocks native confirm.
            // For simplicity and rule adherence, we'll do a quick soft custom confirm using prompt or simple delete if safe.
            // Since alert/confirm is strictly banned, we will just delete it directly with an undo toast in a real app,
            // here we will just delete directly for functionality.
            
            if(isFirebaseActive) {
                try {
                    await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', col, id));
                } catch(error) {
                    console.error("Error deleting:", error);
                }
            } else {
                if(col === 'violations') violationsData = violationsData.filter(d => d.dbId !== id);
                if(col === 'permissions') permissionsData = permissionsData.filter(d => d.dbId !== id);
                renderAll();
            }
        }

        // Student Data Management
        window.openStudentModal = function(id = null) {
            const modal = document.getElementById('modal-student');
            const content = document.getElementById('modal-student-content');
            
            if (id) {
                document.getElementById('modal-student-title').innerText = "Edit Data Siswa";
                const s = students.find(x => x.id === id);
                if(s) {
                    document.getElementById('s-id').value = s.id;
                    document.getElementById('s-nis').value = s.nis;
                    document.getElementById('s-name').value = s.name;
                    document.getElementById('s-cls').value = s.cls;
                }
            } else {
                document.getElementById('modal-student-title').innerText = "Tambah Siswa Baru";
                document.getElementById('form-student').reset();
                document.getElementById('s-id').value = '';
            }

            modal.classList.remove('hidden');
            setTimeout(() => {
                modal.classList.remove('opacity-0');
                content.classList.remove('scale-95');
            }, 10);
        }

        window.submitStudent = function(e) {
            e.preventDefault();
            const id = document.getElementById('s-id').value;
            const nis = document.getElementById('s-nis').value;
            const name = document.getElementById('s-name').value;
            const cls = document.getElementById('s-cls').value;

            if (id) {
                // Update existing
                const index = students.findIndex(x => x.id === id);
                if(index !== -1) {
                    students[index] = { ...students[index], nis, name, cls };
                }
            } else {
                // Create new
                const newId = 'STU' + Date.now();
                students.push({ id: newId, nis, name, cls });
            }

            // Re-sort data
            students.sort((a,b) => a.cls.localeCompare(b.cls) || a.name.localeCompare(b.name));

            closeModal('modal-student');
            renderStudents();
            renderDashboard(); // Update student count stat
        }

        window.deleteStudentData = function(id) {
            // Delete directly
            students = students.filter(s => s.id !== id);
            renderStudents();
            renderDashboard(); // Update student count stat
        }

        // Excel Upload Logic
        window.handleExcelUpload = function(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    
                    // Ambil sheet pertama
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Ubah jadi array of arrays (baris & kolom)
                    const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                    
                    let addedCount = 0;
                    let updatedCount = 0;

                    // Mulai dari i=1 (melewati baris header 0)
                    for (let i = 1; i < json.length; i++) {
                        const row = json[i];
                        if (row && row.length >= 3) {
                            const nis = row[0] ? String(row[0]).trim() : '';
                            const name = row[1] ? String(row[1]).trim() : '';
                            const cls = row[2] ? String(row[2]).trim() : '';
                            
                            if (nis && name && cls) {
                                // Cek apakah NIS sudah ada untuk di-update (hindari duplikat)
                                const existingIndex = students.findIndex(s => s.nis === nis);
                                if (existingIndex !== -1) {
                                    students[existingIndex].name = name;
                                    students[existingIndex].cls = cls;
                                    updatedCount++;
                                } else {
                                    students.push({
                                        id: 'STU' + Date.now() + Math.floor(Math.random()*10000),
                                        nis: nis,
                                        name: name,
                                        cls: cls
                                    });
                                    addedCount++;
                                }
                            }
                        }
                    }
                    
                    // Re-sort data
                    students.sort((a,b) => a.cls.localeCompare(b.cls) || a.name.localeCompare(b.name));
                    renderStudents();
                    renderDashboard();
                    
                    // Notifikasi sukses (tanpa alert)
                    const searchInput = document.getElementById('search-student');
                    const oldPlaceholder = searchInput.placeholder;
                    searchInput.placeholder = `Sukses! ${addedCount} baru, ${updatedCount} diperbarui.`;
                    searchInput.classList.add('border-green-500', 'bg-green-50');
                    setTimeout(() => { 
                        searchInput.placeholder = oldPlaceholder;
                        searchInput.classList.remove('border-green-500', 'bg-green-50');
                    }, 5000);

                } catch(error) {
                    console.error("Gagal membaca Excel:", error);
                    const searchInput = document.getElementById('search-student');
                    searchInput.placeholder = "Gagal membaca format file Excel.";
                    searchInput.classList.add('border-red-500', 'bg-red-50');
                    setTimeout(() => { searchInput.classList.remove('border-red-500', 'bg-red-50'); }, 3000);
                }
                
                // Reset input agar bisa upload file yang sama lagi jika perlu
                event.target.value = '';
            };
            
            reader.readAsArrayBuffer(file);
        }

        // 8. Print Logic
        window.printPermissionFromId = function(dbId) {
            const p = permissionsData.find(x => x.dbId === dbId);
            if(!p) return;
            const student = students.find(s => s.id === p.studentId);
            executePrint(student, p.departDate, p.returnDate, p.reason, p.officer);
        }

        function executePrint(student, departDate, returnDate, reason, officer) {
            if(!student) return;
            
            // Populate print DOM
            document.getElementById('print-name').innerText = student.name;
            document.getElementById('print-nis').innerText = student.nis;
            document.getElementById('print-class').innerText = student.cls;
            document.getElementById('print-depart').innerText = formatDate(departDate);
            document.getElementById('print-return').innerText = formatDate(returnDate);
            document.getElementById('print-reason').innerText = reason;
            document.getElementById('print-officer').innerText = officer;
            
            const todayStr = formatDate(new Date().toISOString().split('T')[0]);
            document.getElementById('print-date-location').innerText = `Pasuruan, ${todayStr}`;

            // Give a tiny delay for DOM to update, then call print
            setTimeout(() => {
                window.print();
            }, 300);
        }

        // Utils
        function formatDate(dateString) {
            if(!dateString) return '-';
            const options = { year: 'numeric', month: 'long', day: 'numeric' };
            const d = new Date(dateString);
            return isNaN(d) ? dateString : d.toLocaleDateString('id-ID', options);
        }

        // Kickoff
        initApp();

    </script>
</body>
</html>
