[index.html.html](https://github.com/user-attachments/files/34188736/index.html.html)
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SARPRAS - SRMP 1 Kabupaten Jepara </title>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <!-- Vue 3 -->
    <script src="https://unpkg.com/vue@3/dist/vue.global.js"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS (XLSX) for Excel Import & Template Download -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: { sans: ['Inter', 'sans-serif'] },
                    colors: {
                        kemensos: {
                            50: '#fff1f2',
                            100: '#ffe4e6',
                            500: '#e11d48',
                            600: '#be123c',
                            700: '#9f1239',
                            800: '#881337',
                            900: '#4c0519'
                        }
                    }
                }
            }
        }
    </script>

    <style>
        [v-cloak] { display: none !important; }
        @media print {
            .no-print { display: none !important; }
            .print-only { display: block !important; }
            body { background: white !important; font-size: 11pt; color: black !important; }
            .print-card { border: none !important; shadow: none !important; padding: 0 !important; }
            .fixed { position: static !important; background: white !important; inset: auto !important; }
            .shadow-2xl, .shadow-xl, .shadow-md, .shadow-sm { box-shadow: none !important; }
            .max-w-3xl, .max-w-4xl { max-width: 100% !important; }
        }
        .print-only { display: none; }
    </style>
</head>
<body class="bg-slate-100 font-sans text-slate-800 antialiased min-h-screen">

    <div id="app" v-cloak class="min-h-screen flex flex-col md:flex-row">

        <!-- SIDEBAR NAVIGATION -->
        <aside class="w-full md:w-64 bg-slate-900 text-white flex-shrink-0 no-print flex flex-col justify-between shadow-xl">
            <div>
                <!-- Brand Header -->
                <div class="p-4 border-b border-slate-800 flex items-center gap-3 bg-slate-950">
                    <img v-if="schoolProfile.logoUrl" :src="schoolProfile.logoUrl" class="w-10 h-10 rounded-xl object-cover bg-white shadow-md shrink-0">
                    <div v-else class="w-10 h-10 rounded-xl bg-kemensos-600 flex items-center justify-center text-white shadow-md font-extrabold text-lg shrink-0">
                        <i class="fas fa-hands-holding-child"></i>
                    </div>
                    <div class="overflow-hidden">
                        <h1 class="font-extrabold text-xs text-white truncate uppercase tracking-wide" :title="schoolProfile.name">{{ schoolProfile.name }}</h1>
                        <p class="text-[9px] text-kemensos-400 font-bold uppercase tracking-wider truncate">KEMENSOS REPUBLIK INDONESIA</p>
                    </div>
                </div>

                <!-- Profile Info Mini Banner -->
                <div class="p-3.5 bg-slate-800/60 border-b border-slate-800 flex items-center justify-between">
                    <div class="truncate pr-2">
                        <p class="text-[11px] font-semibold text-slate-300 truncate">{{ schoolProfile.satker }}</p>
                    </div>
                    <button @click="openSchoolModal" title="Pengaturan Sekolah Rakyat" class="p-1.5 rounded-lg bg-slate-700 hover:bg-kemensos-600 text-slate-300 hover:text-white transition shrink-0">
                        <i class="fas fa-sliders text-xs"></i>
                    </button>
                </div>

                <!-- Nav Menu -->
                <nav class="p-3 space-y-1">
                    <button @click="currentTab = 'dashboard'" :class="currentTab === 'dashboard' ? 'bg-kemensos-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-800 hover:text-white'" class="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl font-medium text-xs transition">
                        <i class="fas fa-chart-pie w-4"></i> Dashboard
                    </button>
                    <button @click="currentTab = 'inventory'" :class="currentTab === 'inventory' ? 'bg-kemensos-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-800 hover:text-white'" class="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl font-medium text-xs transition">
                        <i class="fas fa-boxes-stacked w-4"></i> Data Inventaris Barang
                    </button>
                    <button @click="currentTab = 'borrowings'" :class="currentTab === 'borrowings' ? 'bg-kemensos-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-800 hover:text-white'" class="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl font-medium text-xs transition">
                        <i class="fas fa-hand-holding-hand w-4"></i> Pinjam Pakai Peralatan
                    </button>
                    <button @click="currentTab = 'maintenance'" :class="currentTab === 'maintenance' ? 'bg-kemensos-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-800 hover:text-white'" class="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl font-medium text-xs transition">
                        <i class="fas fa-screwdriver-wrench w-4"></i> Laporan Kerusakan
                    </button>
                    <button @click="currentTab = 'rooms'" :class="currentTab === 'rooms' ? 'bg-kemensos-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-800 hover:text-white'" class="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl font-medium text-xs transition">
                        <i class="fas fa-door-open w-4"></i> Data Ruangan
                    </button>
                    <button @click="currentTab = 'report'" :class="currentTab === 'report' ? 'bg-kemensos-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-800 hover:text-white'" class="w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl font-medium text-xs transition">
                        <i class="fas fa-file-invoice w-4"></i> Cetak Laporan Resmi
                    </button>
                </nav>
            </div>

            <!-- Footer / Clear Storage Button -->
            <div class="p-3 border-t border-slate-800 space-y-2">
                <button @click="openBackupModal" class="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-[11px] font-semibold transition">
                    <i class="fas fa-floppy-disk text-kemensos-400"></i> Simpan & Backup Data
                </button>
                <button @click="openClearModal" class="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-xl bg-slate-800 hover:bg-rose-900/40 text-rose-300 border border-rose-900/30 text-[11px] font-semibold transition">
                    <i class="fas fa-trash-can"></i> Kosongkan / Reset Data
                </button>
            </div>
        </aside>

        <!-- MAIN CONTENT AREA -->
        <main class="flex-1 flex flex-col min-w-0 overflow-hidden">

            <!-- TOP NAVBAR -->
            <header class="bg-white border-b border-slate-200 px-6 py-3 flex items-center justify-between no-print shadow-sm">
                <div class="flex items-center gap-3 min-w-0">
                    <img v-if="schoolProfile.logoUrl" :src="schoolProfile.logoUrl" class="w-9 h-9 rounded-xl object-cover bg-white shadow-sm border border-slate-200 shrink-0">
                    <div v-else class="w-9 h-9 rounded-xl bg-kemensos-600 text-white flex items-center justify-center shrink-0 shadow-sm font-bold text-sm">
                        <i class="fas fa-hands-holding-child"></i>
                    </div>
                    <div class="min-w-0">
                        <div class="flex items-center gap-2">
                            <h2 class="text-sm font-black text-slate-900 truncate leading-tight">{{ schoolProfile.name }}</h2>
                            <span class="px-2 py-0.5 bg-kemensos-100 text-kemensos-700 text-[9px] font-extrabold uppercase rounded tracking-wider shrink-0 hidden sm:inline-block">SARPRAS</span>
                        </div>
                        <p class="text-[11px] text-slate-500 font-medium truncate mt-0.5">
                            <span v-if="currentTab === 'dashboard'">Dashboard Ringkasan Aset</span>
                            <span v-if="currentTab === 'inventory'">Pencatatan Inventaris Barang</span>
                            <span v-if="currentTab === 'borrowings'">Layanan Pinjam Pakai</span>
                            <span v-if="currentTab === 'maintenance'">Pemeliharaan & Perbaikan</span>
                            <span v-if="currentTab === 'rooms'">Data Lokasi & Ruangan</span>
                            <span v-if="currentTab === 'report'">Cetak Laporan Pertanggungjawaban</span>
                        </p>
                    </div>
                </div>

                <!-- User Profile Dropdown -->
                <div class="relative shrink-0">
                    <button @click="showUserDropdown = !showUserDropdown" class="flex items-center gap-2.5 p-1.5 hover:bg-slate-100 rounded-xl transition">
                        <div class="w-8 h-8 rounded-lg bg-slate-900 text-white font-bold flex items-center justify-center text-xs shadow-sm">
                            {{ userProfile.avatarText }}
                        </div>
                        <div class="text-left hidden sm:block">
                            <p class="text-xs font-bold text-slate-800 leading-tight">{{ userProfile.name }}</p>
                            <p class="text-[10px] text-slate-500 font-medium">{{ userProfile.role }}</p>
                        </div>
                        <i class="fas fa-chevron-down text-slate-400 text-[10px]"></i>
                    </button>

                    <!-- User Menu Dropdown Modal -->
                    <div v-if="showUserDropdown" class="absolute right-0 mt-2 w-56 bg-white rounded-2xl shadow-xl border border-slate-100 py-1.5 z-50">
                        <div class="px-4 py-2 border-b border-slate-100">
                            <p class="text-xs font-bold text-slate-800">{{ userProfile.name }}</p>
                            <p class="text-[10px] text-slate-500 truncate">{{ userProfile.email }}</p>
                        </div>
                        <button @click="openAccountModal" class="w-full text-left px-4 py-2 text-xs text-slate-700 hover:bg-slate-50 flex items-center gap-2">
                            <i class="fas fa-user-gear text-slate-400"></i> Edit Profil Operator
                        </button>
                        <button @click="openSchoolModal" class="w-full text-left px-4 py-2 text-xs text-slate-700 hover:bg-slate-50 flex items-center gap-2">
                            <i class="fas fa-landmark text-slate-400"></i> Identitas Sekolah Rakyat
                        </button>
                        <button @click="openBackupModal" class="w-full text-left px-4 py-2 text-xs text-slate-700 hover:bg-slate-50 flex items-center gap-2 border-t border-slate-100">
                            <i class="fas fa-database text-kemensos-600"></i> Backup & Restore Data
                        </button>
                    </div>
                </div>
            </header>

            <!-- TAB CONTENT -->
            <div class="flex-1 overflow-y-auto p-4 md:p-6">

                <!-- 1. DASHBOARD TAB -->
                <div v-if="currentTab === 'dashboard'" class="space-y-6">
                    <!-- Banner Info -->
                    <div class="bg-gradient-to-r from-slate-900 via-slate-800 to-kemensos-900 rounded-2xl p-6 text-white shadow-lg flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                        <div>
                            <span class="inline-block px-3 py-1 bg-kemensos-600/60 backdrop-blur text-[10px] font-bold uppercase rounded-full tracking-wider mb-2 text-kemensos-100">Sistem SARPRAS Terpadu</span>
                            <h3 class="text-xl font-extrabold">{{ schoolProfile.name }}</h3>
                            <p class="text-xs text-slate-300 mt-1 max-w-2xl">{{ schoolProfile.address }} | Satker: {{ schoolProfile.satker }}</p>
                        </div>
                        <button @click="openAddModal" class="px-4 py-2.5 bg-kemensos-600 hover:bg-kemensos-500 text-white text-xs font-bold rounded-xl shadow-md transition flex items-center gap-2 shrink-0">
                            <i class="fas fa-plus"></i> Tambah Aset Baru
                        </button>
                    </div>

                    <!-- Stat Cards Grid -->
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
                        <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex items-center justify-between">
                            <div>
                                <p class="text-[11px] font-semibold text-slate-500 uppercase">Total Unit Barang</p>
                                <h4 class="text-2xl font-black text-slate-800 mt-1">{{ totalItemsCount }}</h4>
                            </div>
                            <div class="w-11 h-11 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center text-lg">
                                <i class="fas fa-box-open"></i>
                            </div>
                        </div>
                        <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex items-center justify-between">
                            <div>
                                <p class="text-[11px] font-semibold text-slate-500 uppercase">Kondisi Baik</p>
                                <h4 class="text-2xl font-black text-emerald-600 mt-1">{{ goodConditionCount }}</h4>
                            </div>
                            <div class="w-11 h-11 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center text-lg">
                                <i class="fas fa-circle-check"></i>
                            </div>
                        </div>
                        <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex items-center justify-between">
                            <div>
                                <p class="text-[11px] font-semibold text-slate-500 uppercase">Rusak / Perlu Perbaikan</p>
                                <h4 class="text-2xl font-black text-rose-600 mt-1">{{ damagedCount }}</h4>
                            </div>
                            <div class="w-11 h-11 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center text-lg">
                                <i class="fas fa-triangle-exclamation"></i>
                            </div>
                        </div>
                        <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex items-center justify-between">
                            <div>
                                <p class="text-[11px] font-semibold text-slate-500 uppercase">Sedang Dipinjam</p>
                                <h4 class="text-2xl font-black text-amber-600 mt-1">{{ activeBorrowingsCount }}</h4>
                            </div>
                            <div class="w-11 h-11 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center text-lg">
                                <i class="fas fa-hand-holding"></i>
                            </div>
                        </div>
                    </div>

                    <!-- Charts Grid -->
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div class="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm">
                            <h4 class="font-bold text-slate-800 text-sm mb-4 flex items-center gap-2">
                                <i class="fas fa-chart-donut text-kemensos-600"></i>
                                Proporsi Kondisi Aset Barang
                            </h4>
                            <div v-if="inventory.length > 0" class="h-60 relative">
                                <canvas id="conditionChart"></canvas>
                            </div>
                            <div v-else class="h-60 flex flex-col items-center justify-center text-center text-slate-400 p-4">
                                <i class="fas fa-inbox text-4xl mb-2 text-slate-300"></i>
                                <p class="text-xs">Belum ada data barang untuk ditampilkan</p>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm">
                            <h4 class="font-bold text-slate-800 text-sm mb-4 flex items-center gap-2">
                                <i class="fas fa-chart-bar text-kemensos-600"></i>
                                Distribusi Per Kategori
                            </h4>
                            <div v-if="inventory.length > 0" class="h-60 relative">
                                <canvas id="categoryChart"></canvas>
                            </div>
                            <div v-else class="h-60 flex flex-col items-center justify-center text-center text-slate-400 p-4">
                                <i class="fas fa-folder-open text-4xl mb-2 text-slate-300"></i>
                                <p class="text-xs">Belum ada data ketersediaan kategori</p>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 2. INVENTORY TAB -->
                <div v-if="currentTab === 'inventory'" class="space-y-4">
                    <!-- Action Bar & Filters -->
                    <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex flex-col md:flex-row gap-3 justify-between items-center">
                        <div class="flex flex-wrap items-center gap-2 w-full md:w-auto">
                            <div class="relative flex-1 md:w-64">
                                <i class="fas fa-search absolute left-3 top-3 text-slate-400 text-xs"></i>
                                <input v-model="invSearch" type="text" placeholder="Cari nama / kode barang..." class="w-full pl-9 pr-3 py-2 text-xs border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                            </div>
                            <select v-model="invCategoryFilter" class="px-3 py-2 text-xs border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                                <option value="">Semua Kategori</option>
                                <option v-for="c in categories" :key="c" :value="c">{{ c }}</option>
                            </select>
                            <select v-model="invConditionFilter" class="px-3 py-2 text-xs border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                                <option value="">Semua Kondisi</option>
                                <option value="Baik">Baik</option>
                                <option value="Rusak Ringan">Rusak Ringan</option>
                                <option value="Rusak Berat">Rusak Berat</option>
                            </select>
                        </div>
                        <div class="flex items-center gap-2 w-full md:w-auto justify-end shrink-0">
                            <button @click="openImportExcelModal" class="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold rounded-xl shadow transition flex items-center gap-1.5">
                                <i class="fas fa-file-excel"></i> Import / Template Excel
                            </button>
                            <button @click="openCategoryModal" class="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold rounded-xl shadow transition flex items-center gap-1.5">
                                <i class="fas fa-tags"></i> Kelola Kategori
                            </button>
                            <button @click="openAddModal" class="px-4 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white text-xs font-semibold rounded-xl shadow transition flex items-center gap-1.5">
                                <i class="fas fa-plus"></i> Tambah Barang
                            </button>
                        </div>
                    </div>

                    <!-- Inventory Table -->
                    <div class="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
                        <div class="overflow-x-auto">
                            <table class="w-full text-left text-xs">
                                <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider">
                                    <tr>
                                        <th class="py-3.5 px-4">Kode & Nama Barang</th>
                                        <th class="py-3.5 px-4">Kategori</th>
                                        <th class="py-3.5 px-4">Lokasi Ruangan</th>
                                        <th class="py-3.5 px-4 text-center">Jumlah</th>
                                        <th class="py-3.5 px-4">Kondisi</th>
                                        <th class="py-3.5 px-4">Sumber Anggaran</th>
                                        <th class="py-3.5 px-4 text-right">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-100">
                                    <tr v-for="item in filteredInventory" :key="item.id" class="hover:bg-slate-50/80 transition">
                                        <td class="py-3 px-4 font-semibold text-slate-800">
                                            <div class="font-bold text-slate-900">{{ item.name }}</div>
                                            <span class="inline-block px-2 py-0.5 bg-slate-100 text-slate-600 rounded text-[10px] font-mono mt-0.5">{{ item.code }}</span>
                                        </td>
                                        <td class="py-3 px-4 text-slate-600">{{ item.category }}</td>
                                        <td class="py-3 px-4 text-slate-600 font-medium">
                                            <i class="fas fa-location-dot text-slate-400 mr-1"></i> {{ item.room }}
                                        </td>
                                        <td class="py-3 px-4 text-center font-bold text-slate-800">
                                            {{ item.qty }} <span class="text-[10px] text-slate-500 font-normal">{{ item.unit }}</span>
                                        </td>
                                        <td class="py-3 px-4">
                                            <span :class="{
                                                'bg-emerald-100 text-emerald-800': item.condition === 'Baik',
                                                'bg-amber-100 text-amber-800': item.condition === 'Rusak Ringan',
                                                'bg-rose-100 text-rose-800': item.condition === 'Rusak Berat'
                                            }" class="px-2.5 py-1 rounded-full text-[10px] font-bold inline-block">
                                                {{ item.condition }}
                                            </span>
                                        </td>
                                        <td class="py-3 px-4 text-slate-600">
                                            {{ item.fundSource }} <span class="text-[10px] text-slate-400">({{ item.year }})</span>
                                        </td>
                                        <td class="py-3 px-4 text-right space-x-1">
                                            <button @click="editItem(item)" class="p-1.5 text-blue-600 hover:bg-blue-50 rounded-lg transition" title="Edit">
                                                <i class="fas fa-pen-to-square"></i>
                                            </button>
                                            <button @click="confirmDeleteItem(item)" class="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg transition" title="Hapus">
                                                <i class="fas fa-trash-can"></i>
                                            </button>
                                        </td>
                                    </tr>
                                    <tr v-if="filteredInventory.length === 0">
                                        <td colspan="7" class="py-12 text-center text-slate-400">
                                            <i class="fas fa-folder-open text-4xl mb-3 text-slate-300"></i>
                                            <p class="font-medium text-xs">Belum ada data barang terdaftar</p>
                                            <p class="text-[11px] text-slate-400 mt-0.5">Silakan klik tombol "Tambah Barang" untuk mengisi inventaris</p>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <!-- 3. BORROWINGS TAB -->
                <div v-if="currentTab === 'borrowings'" class="space-y-4">
                    <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex justify-between items-center">
                        <div>
                            <h3 class="font-bold text-slate-800 text-sm">Catatan Pinjam Pakai Peralatan</h3>
                            <p class="text-xs text-slate-500">Mencatat peminjaman alat vokasional/elektronik oleh siswa, tutor, atau pendamping sosial.</p>
                        </div>
                        <button @click="openBorrowModal" class="px-4 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white text-xs font-semibold rounded-xl shadow transition flex items-center gap-1.5">
                            <i class="fas fa-plus"></i> Pinjamkan Barang
                        </button>
                    </div>

                    <div class="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
                        <table class="w-full text-left text-xs">
                            <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase">
                                <tr>
                                    <th class="py-3.5 px-4">Peminjam</th>
                                    <th class="py-3.5 px-4">Daftar Barang Dipinjam</th>
                                    <th class="py-3.5 px-4">Tanggal Pinjam</th>
                                    <th class="py-3.5 px-4">Rencana Kembali</th>
                                    <th class="py-3.5 px-4">Status</th>
                                    <th class="py-3.5 px-4 text-right">Aksi</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100">
                                <tr v-for="b in borrowings" :key="b.id" class="hover:bg-slate-50 transition">
                                    <td class="py-3 px-4 font-semibold text-slate-800">
                                        {{ b.borrowerName }}
                                        <span class="block text-[10px] text-slate-500 font-normal">{{ b.role }}</span>
                                    </td>
                                    <td class="py-3 px-4 text-slate-700 font-medium">
                                        <ul class="space-y-1">
                                            <li v-for="(itemRow, idx) in getBorrowingItemsList(b)" :key="idx" class="flex items-center gap-1.5">
                                                <span class="w-1.5 h-1.5 rounded-full bg-kemensos-600 shrink-0"></span>
                                                <span class="font-bold text-slate-900">{{ itemRow.name }}</span>
                                                <span class="text-[10px] px-1.5 py-0.5 bg-slate-100 rounded text-slate-600 font-semibold">({{ itemRow.qtyBorrowed }} {{ itemRow.unit }})</span>
                                            </li>
                                        </ul>
                                    </td>
                                    <td class="py-3 px-4 text-slate-600">{{ formatDate(b.borrowDate) }}</td>
                                    <td class="py-3 px-4 text-amber-700 font-medium">{{ formatDate(b.returnPlanDate) }}</td>
                                    <td class="py-3 px-4">
                                        <span :class="b.status === 'Dipinjam' ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800'" class="px-2.5 py-1 rounded-full text-[10px] font-bold">
                                            {{ b.status }}
                                        </span>
                                    </td>
                                    <td class="py-3 px-4 text-right space-x-1">
                                        <button v-if="b.status === 'Dipinjam'" @click="returnItem(b)" class="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[11px] font-bold shadow-sm transition">
                                            Kembalikan
                                        </button>
                                        <span v-else class="text-[11px] text-slate-400 font-medium mr-1">Selesai</span>
                                        <button @click="openBorrowLetter(b)" class="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg transition" title="Cetak Surat Peminjaman / Berkas">
                                            <i class="fas fa-file-pdf"></i>
                                        </button>
                                        <button @click="deleteBorrowing(b.id)" class="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg transition" title="Hapus Catatan">
                                            <i class="fas fa-trash-can"></i>
                                        </button>
                                    </td>
                                </tr>
                                <tr v-if="borrowings.length === 0">
                                    <td colspan="6" class="py-12 text-center text-slate-400">
                                        <i class="fas fa-hand-holding text-4xl mb-2 text-slate-300"></i>
                                        <p class="text-xs">Belum ada catatan peminjaman aktif</p>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- 4. MAINTENANCE TAB -->
                <div v-if="currentTab === 'maintenance'" class="space-y-4">
                    <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex justify-between items-center">
                        <div>
                            <h3 class="font-bold text-slate-800 text-sm">Laporan Kerusakan & Pemeliharaan</h3>
                            <p class="text-xs text-slate-500">Mencatat barang yang membutuhkan perbaikan teknis atau pengusulan servis.</p>
                        </div>
                        <button @click="openMaintenanceModal" class="px-4 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white text-xs font-semibold rounded-xl shadow transition flex items-center gap-1.5">
                            <i class="fas fa-plus"></i> Laporkan Kerusakan
                        </button>
                    </div>

                    <div class="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
                        <table class="w-full text-left text-xs">
                            <thead class="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold uppercase">
                                <tr>
                                    <th class="py-3.5 px-4">Barang</th>
                                    <th class="py-3.5 px-4">Deskripsi Kerusakan</th>
                                    <th class="py-3.5 px-4">Pelapor</th>
                                    <th class="py-3.5 px-4">Est. Biaya</th>
                                    <th class="py-3.5 px-4">Status Perbaikan</th>
                                    <th class="py-3.5 px-4 text-right">Aksi Status</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100">
                                <tr v-for="m in maintenances" :key="m.id" class="hover:bg-slate-50 transition">
                                    <td class="py-3 px-4 font-bold text-slate-800">{{ getItemName(m.itemId) }}</td>
                                    <td class="py-3 px-4 text-slate-600 max-w-xs">{{ m.description }}</td>
                                    <td class="py-3 px-4 text-slate-600">{{ m.reporter }}</td>
                                    <td class="py-3 px-4 font-semibold text-slate-800">{{ formatCurrency(m.cost) }}</td>
                                    <td class="py-3 px-4">
                                        <span :class="{
                                            'bg-rose-100 text-rose-800': m.status === 'Menunggu',
                                            'bg-amber-100 text-amber-800': m.status === 'Proses Perbaikan',
                                            'bg-emerald-100 text-emerald-800': m.status === 'Selesai'
                                        }" class="px-2.5 py-1 rounded-full text-[10px] font-bold">
                                            {{ m.status }}
                                        </span>
                                    </td>
                                    <td class="py-3 px-4 text-right space-x-1">
                                        <button v-if="m.status === 'Menunggu'" @click="updateMaintenanceStatus(m, 'Proses Perbaikan')" class="px-2.5 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded text-[10px] font-bold">
                                            Proses
                                        </button>
                                        <button v-if="m.status === 'Proses Perbaikan'" @click="updateMaintenanceStatus(m, 'Selesai')" class="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[10px] font-bold">
                                            Selesaikan
                                        </button>
                                        <button @click="deleteMaintenance(m.id)" class="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg transition" title="Hapus Laporan">
                                            <i class="fas fa-trash-can"></i>
                                        </button>
                                    </td>
                                </tr>
                                <tr v-if="maintenances.length === 0">
                                    <td colspan="6" class="py-12 text-center text-slate-400">
                                        <i class="fas fa-screwdriver-wrench text-4xl mb-2 text-slate-300"></i>
                                        <p class="text-xs">Tidak ada usulan perbaikan saat ini</p>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- 5. ROOMS TAB -->
                <div v-if="currentTab === 'rooms'" class="space-y-4">
                    <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex justify-between items-center">
                        <div>
                            <h3 class="font-bold text-slate-800 text-sm">Daftar Ruangan Sekolah Rakyat</h3>
                            <p class="text-xs text-slate-500">Pemetaan ruangan kelas, bengkel pelatihan vokasi, kantor, dan fasilitas pendukung.</p>
                        </div>
                        <button @click="openRoomModal" class="px-4 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white text-xs font-semibold rounded-xl shadow transition flex items-center gap-1.5">
                            <i class="fas fa-plus"></i> Tambah Ruangan
                        </button>
                    </div>

                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div v-for="r in roomBreakdown" :key="r.id" class="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm hover:border-kemensos-500 transition">
                            <div class="flex justify-between items-start mb-3">
                                <div class="w-10 h-10 rounded-xl bg-slate-100 text-kemensos-600 flex items-center justify-center text-lg font-bold">
                                    <i class="fas fa-door-closed"></i>
                                </div>
                                <div class="flex items-center gap-1">
                                    <span class="px-2.5 py-1 bg-slate-100 text-slate-700 rounded-lg text-[10px] font-bold">
                                        {{ r.totalUnits }} Unit Total
                                    </span>
                                    <button @click="editRoom(r)" class="p-1 text-slate-400 hover:text-blue-600 rounded transition" title="Edit Ruangan">
                                        <i class="fas fa-pen-to-square text-xs"></i>
                                    </button>
                                    <button @click="confirmDeleteRoom(r)" class="p-1 text-slate-400 hover:text-rose-600 rounded transition" title="Hapus Ruangan">
                                        <i class="fas fa-trash-can text-xs"></i>
                                    </button>
                                </div>
                            </div>
                            <h4 class="font-bold text-slate-800 text-sm">{{ r.name }}</h4>
                            <p class="text-xs text-slate-500 mt-1">PIC / Penanggung Jawab: <span class="font-semibold text-slate-700">{{ r.pic || '-' }}</span></p>
                            <div class="mt-4 pt-3 border-t border-slate-100 flex justify-between text-[11px] text-slate-500">
                                <span>Jenis Item: <strong>{{ r.itemTypesCount }} Tipe</strong></span>
                                <button @click="filterByRoom(r.name)" class="text-kemensos-600 font-bold hover:underline">Lihat Barang &rarr;</button>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- 6. REPORT & PRINT TAB -->
                <div v-if="currentTab === 'report'" class="space-y-6">
                    <!-- Top Action Bar -->
                    <div class="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex justify-between items-center no-print">
                        <div>
                            <h3 class="font-bold text-slate-800 text-sm">Laporan Pertanggungjawaban SARPRAS</h3>
                            <p class="text-xs text-slate-500">Dokumen resmi berformat KOP Kementerian Sosial RI untuk diarsip atau dicetak.</p>
                        </div>
                        <button @click="printReport" class="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold rounded-xl shadow-md transition flex items-center gap-2">
                            <i class="fas fa-print"></i> Cetak Dokumen / PDF
                        </button>
                    </div>

                    <!-- PRINT PREVIEW SHEET -->
                    <div class="bg-white rounded-2xl p-8 border border-slate-200 shadow-sm max-w-4xl mx-auto print-card">
                        
                        <!-- OFFICIAL KOP SURAT KEMENSOS -->
                        <div class="border-b-4 border-double border-slate-900 pb-4 mb-6 flex items-center justify-between text-center">
                            <div class="w-16 h-16 flex items-center justify-center shrink-0">
                                <img v-if="schoolProfile.logoUrl" :src="schoolProfile.logoUrl" class="max-w-full max-h-full object-contain">
                                <i v-else class="fas fa-hands-holding-child text-4xl text-kemensos-700"></i>
                            </div>
                            <div class="flex-1 px-4">
                                <h2 class="text-sm font-bold tracking-wider text-slate-900 uppercase">KEMENTERIAN SOSIAL REPUBLIK INDONESIA</h2>
                                <h1 class="text-base font-black tracking-wide text-slate-900 uppercase mt-0.5">{{ schoolProfile.name }}</h1>
                                <p class="text-[10px] text-slate-600 mt-0.5">{{ schoolProfile.address }}</p>
                                <p class="text-[10px] text-slate-600">Email: {{ schoolProfile.email }}</p>
                            </div>
                            <div class="w-16"></div>
                        </div>

                        <!-- REPORT TITLE -->
                        <div class="text-center mb-6">
                            <h3 class="text-sm font-bold uppercase underline tracking-wide text-slate-900">BERITA ACARA REKAPITULASI SARANA DAN PRASARANA</h3>
                            <p class="text-xs text-slate-600 mt-1">Nomor Lampiran: SARPRAS-SRMP/LAP/{{ new Date().getFullYear() }}</p>
                        </div>

                        <!-- INVENTORY SUMMARY TABLE FOR PRINT -->
                        <table class="w-full text-left text-xs border-collapse border border-slate-300 mb-8">
                            <thead>
                                <tr class="bg-slate-100 text-slate-800 font-bold border-b border-slate-300">
                                    <th class="p-2 border border-slate-300 text-center w-10">No</th>
                                    <th class="p-2 border border-slate-300">Kode & Nama Barang</th>
                                    <th class="p-2 border border-slate-300">Kategori</th>
                                    <th class="p-2 border border-slate-300">Lokasi</th>
                                    <th class="p-2 border border-slate-300 text-center">Qty</th>
                                    <th class="p-2 border border-slate-300">Kondisi</th>
                                    <th class="p-2 border border-slate-300">Sumber Anggaran</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr v-for="(item, index) in inventory" :key="item.id" class="border-b border-slate-200">
                                    <td class="p-2 border border-slate-300 text-center">{{ index + 1 }}</td>
                                    <td class="p-2 border border-slate-300">
                                        <div class="font-bold">{{ item.name }}</div>
                                        <div class="text-[10px] text-slate-500 font-mono">{{ item.code }}</div>
                                    </td>
                                    <td class="p-2 border border-slate-300">{{ item.category }}</td>
                                    <td class="p-2 border border-slate-300">{{ item.room }}</td>
                                    <td class="p-2 border border-slate-300 text-center font-bold">{{ item.qty }} {{ item.unit }}</td>
                                    <td class="p-2 border border-slate-300">{{ item.condition }}</td>
                                    <td class="p-2 border border-slate-300">{{ item.fundSource }} ({{ item.year }})</td>
                                </tr>
                                <tr v-if="inventory.length === 0">
                                    <td colspan="7" class="p-4 text-center text-slate-400 italic">Data barang masih kosong</td>
                                </tr>
                            </tbody>
                        </table>

                        <!-- SIGNATURE BLOCK -->
                        <div class="grid grid-cols-2 gap-8 text-center text-xs mt-12 pt-6">
                            <div>
                                <p class="mb-1">Mengetahui,</p>
                                <p class="font-bold uppercase text-slate-800">Penanggung Jawab / Kepala Sentra</p>
                                <div class="h-20"></div>
                                <p class="font-bold underline text-slate-900">{{ schoolProfile.principalName }}</p>
                                <p class="text-[10px] text-slate-600">NIP. {{ schoolProfile.principalNip }}</p>
                            </div>
                            <div>
                                <p class="mb-1">{{ schoolProfile.city || 'Kota' }}, {{ formatDate(new Date()) }}</p>
                                <p class="font-bold uppercase text-slate-800">Pengelola SARPRAS</p>
                                <div class="h-20"></div>
                                <p class="font-bold underline text-slate-900">{{ schoolProfile.sarprasHeadName }}</p>
                                <p class="text-[10px] text-slate-600">NIP. {{ schoolProfile.sarprasHeadNip }}</p>
                            </div>
                        </div>

                    </div>
                </div>

            </div>
        </main>

        <!-- ================= MODALS ================= -->

        <!-- MODAL: IMPORT EXCEL & DOWNLOAD TEMPLATE -->
        <div v-if="showImportExcelModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-3xl w-full overflow-hidden flex flex-col max-h-[90vh]">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between shrink-0">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-file-excel text-emerald-600 text-base"></i>
                        <span>Import Data Barang & Download Template Excel</span>
                    </h3>
                    <button @click="showImportExcelModal = false" class="text-slate-400 hover:text-slate-600 rounded-lg p-1 transition">
                        <i class="fas fa-xmark text-base"></i>
                    </button>
                </div>

                <div class="p-6 space-y-5 overflow-y-auto text-xs">
                    <!-- Step 1: Download Template -->
                    <div class="p-4 bg-emerald-50/60 border border-emerald-200 rounded-2xl space-y-3">
                        <div class="flex items-center justify-between flex-wrap gap-2">
                            <div>
                                <h4 class="font-bold text-emerald-900 text-xs flex items-center gap-1.5">
                                    <i class="fas fa-download text-emerald-600"></i> Langkah 1: Download Template Excel
                                </h4>
                                <p class="text-[11px] text-emerald-800 mt-0.5">Gunakan format file Excel resmi yang sudah dikotak-kotakkan dan ditata sesuai kolom aplikasi.</p>
                            </div>
                            <button @click="downloadExcelTemplate" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl shadow text-xs flex items-center gap-2 transition">
                                <i class="fas fa-file-arrow-down"></i> Download Template (.xlsx)
                            </button>
                        </div>

                        <!-- Template Grid Box Preview -->
                        <div class="bg-white rounded-xl border border-emerald-300 overflow-hidden shadow-sm mt-2">
                            <div class="bg-emerald-100/70 text-emerald-900 text-[10px] font-extrabold px-3 py-1.5 uppercase tracking-wider border-b border-emerald-200 flex justify-between items-center">
                                <span><i class="fas fa-table mr-1"></i> Format Box / Kolom Template Excel</span>
                                <span class="text-[9px] font-normal text-emerald-800">Harap ikuti tata letak kolom di bawah ini</span>
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-[11px] text-left border-collapse">
                                    <thead>
                                        <tr class="bg-slate-100 text-slate-700 font-bold border-b border-slate-300">
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50">Kode Barang</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50">Nama Barang</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50">Kategori</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50">Ruangan</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50 text-center">Jumlah</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50">Satuan</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50">Kondisi</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50">Sumber Anggaran</th>
                                            <th class="p-2 border border-slate-300 bg-emerald-50/50 text-center">Tahun</th>
                                        </tr>
                                    </thead>
                                    <tbody class="text-slate-600 italic text-[10px]">
                                        <tr>
                                            <td class="p-1.5 border border-slate-200">SRMP-INV-001</td>
                                            <td class="p-1.5 border border-slate-200 font-bold text-slate-800">Laptop Chromebook Belajar</td>
                                            <td class="p-1.5 border border-slate-200">Perlengkapan Belajar & TIK</td>
                                            <td class="p-1.5 border border-slate-200">Ruang Belajar TIK</td>
                                            <td class="p-1.5 border border-slate-200 text-center font-bold">10</td>
                                            <td class="p-1.5 border border-slate-200">Unit</td>
                                            <td class="p-1.5 border border-slate-200 text-emerald-700 font-bold">Baik</td>
                                            <td class="p-1.5 border border-slate-200">DIPA Kemensos</td>
                                            <td class="p-1.5 border border-slate-200 text-center">2024</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- Step 2: Upload File -->
                    <div class="p-4 border border-slate-200 rounded-2xl space-y-3 bg-slate-50/50">
                        <h4 class="font-bold text-slate-800 text-xs flex items-center gap-1.5">
                            <i class="fas fa-file-upload text-blue-600"></i> Langkah 2: Upload File Excel yang Telah Diisi (.xlsx / .xls / .csv)
                        </h4>
                        <div class="flex items-center gap-3">
                            <input type="file" ref="excelFileInput" @change="handleExcelFileUpload" accept=".xlsx, .xls, .csv" class="w-full text-xs text-slate-600 file:mr-3 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-bold file:bg-blue-600 file:text-white hover:file:bg-blue-700 cursor-pointer border border-slate-300 rounded-xl bg-white">
                        </div>
                    </div>

                    <!-- Step 3: Boxed Preview Table of Excel File Data -->
                    <div v-if="parsedExcelItems.length > 0" class="space-y-3">
                        <div class="flex items-center justify-between">
                            <h4 class="font-bold text-slate-800 text-xs flex items-center gap-2">
                                <i class="fas fa-list-check text-emerald-600"></i> Pratinjau Data File Excel (Terdeteksi {{ parsedExcelItems.length }} Baris)
                            </h4>
                            <span class="text-[10px] bg-blue-100 text-blue-800 font-bold px-2.5 py-0.5 rounded-full">Siap Di-Import</span>
                        </div>

                        <!-- Boxed Preview Table -->
                        <div class="bg-white rounded-xl border border-slate-300 overflow-hidden shadow-sm max-h-60 overflow-y-auto">
                            <table class="w-full text-[11px] text-left border-collapse">
                                <thead>
                                    <tr class="bg-slate-100 text-slate-700 font-bold border-b border-slate-300 sticky top-0 bg-slate-100 z-10">
                                        <th class="p-2 border border-slate-300 text-center w-8">#</th>
                                        <th class="p-2 border border-slate-300">Kode Barang</th>
                                        <th class="p-2 border border-slate-300">Nama Barang</th>
                                        <th class="p-2 border border-slate-300">Kategori</th>
                                        <th class="p-2 border border-slate-300">Ruangan</th>
                                        <th class="p-2 border border-slate-300 text-center">Qty</th>
                                        <th class="p-2 border border-slate-300">Satuan</th>
                                        <th class="p-2 border border-slate-300">Kondisi</th>
                                        <th class="p-2 border border-slate-300">Sumber Anggaran</th>
                                        <th class="p-2 border border-slate-300 text-center">Tahun</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-200 text-slate-700">
                                    <tr v-for="(item, idx) in parsedExcelItems" :key="idx" class="hover:bg-slate-50">
                                        <td class="p-2 border border-slate-200 text-center font-bold text-slate-500">{{ idx + 1 }}</td>
                                        <td class="p-2 border border-slate-200 font-mono text-slate-800">{{ item.code }}</td>
                                        <td class="p-2 border border-slate-200 font-bold text-slate-900">{{ item.name }}</td>
                                        <td class="p-2 border border-slate-200">{{ item.category }}</td>
                                        <td class="p-2 border border-slate-200">{{ item.room }}</td>
                                        <td class="p-2 border border-slate-200 text-center font-bold">{{ item.qty }}</td>
                                        <td class="p-2 border border-slate-200">{{ item.unit }}</td>
                                        <td class="p-2 border border-slate-200">
                                            <span class="px-2 py-0.5 rounded text-[10px] font-bold" :class="{
                                                'bg-emerald-100 text-emerald-800': item.condition === 'Baik',
                                                'bg-amber-100 text-amber-800': item.condition === 'Rusak Ringan',
                                                'bg-rose-100 text-rose-800': item.condition === 'Rusak Berat'
                                            }">{{ item.condition }}</span>
                                        </td>
                                        <td class="p-2 border border-slate-200">{{ item.fundSource }}</td>
                                        <td class="p-2 border border-slate-200 text-center">{{ item.year }}</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>

                        <!-- Import Mode Options -->
                        <div class="flex items-center justify-between pt-2 border-t border-slate-200 flex-wrap gap-2">
                            <div class="flex items-center gap-4">
                                <label class="flex items-center gap-1.5 cursor-pointer text-xs font-semibold text-slate-700">
                                    <input type="radio" v-model="excelImportMode" value="append" class="text-emerald-600 focus:ring-emerald-500">
                                    <span>Gabungkan (Tambahkan ke Data Saat Ini)</span>
                                </label>
                                <label class="flex items-center gap-1.5 cursor-pointer text-xs font-semibold text-rose-700">
                                    <input type="radio" v-model="excelImportMode" value="replace" class="text-rose-600 focus:ring-rose-500">
                                    <span>Ganti Seluruh Data Inventaris</span>
                                </label>
                            </div>
                            <button @click="executeImportExcel" class="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold rounded-xl shadow-md transition flex items-center gap-2">
                                <i class="fas fa-check-circle"></i> Simpan & Import Data
                            </button>
                        </div>
                    </div>
                </div>

                <div class="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end shrink-0">
                    <button @click="showImportExcelModal = false" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold rounded-xl text-xs">Tutup</button>
                </div>
            </div>
        </div>

        <!-- MODAL: SURAT PEMINJAMAN BARANG / BUKTI PINJAM PAKAI -->
        <div v-if="showBorrowLetterModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <div class="bg-white rounded-2xl shadow-2xl max-w-3xl w-full overflow-hidden flex flex-col max-h-[90vh]">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between no-print shrink-0">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-file-signature text-kemensos-600"></i> Cetak Surat Bukti Pinjam Pakai
                    </h3>
                    <div class="flex items-center gap-2">
                        <button @click="printBorrowLetter" class="px-4 py-1.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold rounded-xl shadow transition flex items-center gap-1.5">
                            <i class="fas fa-print"></i> Cetak / Simpan PDF
                        </button>
                        <button @click="showBorrowLetterModal = false" class="text-slate-400 hover:text-slate-600 p-1">
                            <i class="fas fa-xmark text-base"></i>
                        </button>
                    </div>
                </div>

                <div class="p-8 overflow-y-auto print-card text-slate-900 font-sans leading-relaxed">
                    <!-- OFFICIAL KOP SURAT KEMENSOS -->
                    <div class="border-b-4 border-double border-slate-900 pb-4 mb-6 flex items-center justify-between text-center">
                        <div class="w-16 h-16 flex items-center justify-center shrink-0">
                            <img v-if="schoolProfile.logoUrl" :src="schoolProfile.logoUrl" class="max-w-full max-h-full object-contain">
                            <i v-else class="fas fa-hands-holding-child text-4xl text-kemensos-700"></i>
                        </div>
                        <div class="flex-1 px-4">
                            <h2 class="text-xs font-bold tracking-wider text-slate-900 uppercase">KEMENTERIAN SOSIAL REPUBLIK INDONESIA</h2>
                            <h1 class="text-sm font-black tracking-wide text-slate-900 uppercase mt-0.5">{{ schoolProfile.name }}</h1>
                            <p class="text-[10px] text-slate-600 mt-0.5">{{ schoolProfile.address }}</p>
                            <p class="text-[10px] text-slate-600">Email: {{ schoolProfile.email }}</p>
                        </div>
                        <div class="w-16"></div>
                    </div>

                    <!-- TITLE -->
                    <div class="text-center mb-6">
                        <h3 class="text-sm font-bold uppercase underline tracking-wide text-slate-900">SURAT BUKTI PINJAM PAKAI BARANG INVENTARIS</h3>
                        <p class="text-[11px] text-slate-600 mt-0.5">Nomor Register: REG-PINJAM/{{ selectedBorrowing ? selectedBorrowing.id : '00' }}</p>
                    </div>

                    <p class="text-xs mb-3 text-slate-800">Telah diserahterimakan barang inventaris milik unit sekolah/sentra untuk keperluan operasional/pembelajaran dengan rincian sebagai berikut:</p>

                    <!-- IDENTITAS PENERIMA -->
                    <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 mb-4 text-xs space-y-1.5">
                        <div class="font-bold text-slate-800 border-b border-slate-200 pb-1 mb-2 uppercase text-[10px] tracking-wider text-kemensos-700">A. Data Penerima / Peminjam</div>
                        <div class="grid grid-cols-3"><span class="text-slate-500">Nama Peminjam</span><span class="col-span-2 font-bold text-slate-900">: {{ selectedBorrowing?.borrowerName }}</span></div>
                        <div class="grid grid-cols-3"><span class="text-slate-500">Peran / Jabatan</span><span class="col-span-2 font-medium text-slate-800">: {{ selectedBorrowing?.role }}</span></div>
                        <div class="grid grid-cols-3"><span class="text-slate-500">Tanggal Peminjaman</span><span class="col-span-2 font-medium text-slate-800">: {{ formatDate(selectedBorrowing?.borrowDate) }}</span></div>
                        <div class="grid grid-cols-3"><span class="text-slate-500">Rencana Pengembalian</span><span class="col-span-2 font-medium text-amber-800 font-bold">: {{ formatDate(selectedBorrowing?.returnPlanDate) }}</span></div>
                        <div class="grid grid-cols-3"><span class="text-slate-500">Status Saat Ini</span><span class="col-span-2 font-bold" :class="selectedBorrowing?.status === 'Dipinjam' ? 'text-amber-700' : 'text-emerald-700'">: {{ selectedBorrowing?.status }}</span></div>
                    </div>

                    <!-- DETAIL BARANG MULTI ITEM TABLE -->
                    <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 mb-5 text-xs">
                        <div class="font-bold text-slate-800 border-b border-slate-200 pb-1 mb-2 uppercase text-[10px] tracking-wider text-kemensos-700">B. Detail Barang Dipinjam</div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-left text-xs border-collapse">
                                <thead>
                                    <tr class="border-b border-slate-300 text-slate-600 font-bold bg-slate-100">
                                        <th class="p-1.5 text-center w-8">No</th>
                                        <th class="p-1.5">Kode Barang</th>
                                        <th class="p-1.5">Nama Barang</th>
                                        <th class="p-1.5">Kategori</th>
                                        <th class="p-1.5">Lokasi</th>
                                        <th class="p-1.5 text-center">Jumlah</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-200">
                                    <tr v-for="(itemRow, idx) in getBorrowingItemsList(selectedBorrowing)" :key="idx">
                                        <td class="p-1.5 text-center font-bold text-slate-500">{{ idx + 1 }}</td>
                                        <td class="p-1.5 font-mono text-slate-800">{{ itemRow.code }}</td>
                                        <td class="p-1.5 font-bold text-slate-900">{{ itemRow.name }}</td>
                                        <td class="p-1.5 text-slate-600">{{ itemRow.category }}</td>
                                        <td class="p-1.5 text-slate-600">{{ itemRow.room }}</td>
                                        <td class="p-1.5 text-center font-bold text-slate-900">{{ itemRow.qtyBorrowed }} {{ itemRow.unit }}</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- KETENTUAN -->
                    <div class="text-[11px] text-slate-700 bg-amber-50/60 border border-amber-200 p-3.5 rounded-xl mb-8 space-y-1">
                        <p class="font-bold text-amber-900">Ketentuan dan Tanggung Jawab Peminjam:</p>
                        <ol class="list-decimal list-inside space-y-0.5 text-[10px] text-slate-700">
                            <li>Peminjam wajib merawat dan menggunakan barang sesuai peruntukan serta menjaga kebersihan.</li>
                            <li>Apabila terjadi kerusakan atau kehilangan karena kelalaian, peminjam wajib memperbaiki atau mengganti barang.</li>
                            <li>Barang harus segera dikembalikan apabila masa pinjam pakai telah selesai.</li>
                        </ol>
                    </div>

                    <!-- TANDA TANGAN -->
                    <div class="grid grid-cols-3 gap-4 text-center text-xs mt-6">
                        <div>
                            <p class="mb-1 text-slate-600">Peminjam / Penerima,</p>
                            <div class="h-16"></div>
                            <p class="font-bold underline text-slate-900">{{ selectedBorrowing?.borrowerName }}</p>
                            <p class="text-[10px] text-slate-500">({{ selectedBorrowing?.role }})</p>
                        </div>
                        <div>
                            <p class="mb-1 text-slate-600">Pengelola SARPRAS,</p>
                            <div class="h-16"></div>
                            <p class="font-bold underline text-slate-900">{{ schoolProfile.sarprasHeadName }}</p>
                            <p class="text-[10px] text-slate-500">NIP. {{ schoolProfile.sarprasHeadNip }}</p>
                        </div>
                        <div>
                            <p class="mb-1 text-slate-600">Mengetahui Kepala,</p>
                            <div class="h-16"></div>
                            <p class="font-bold underline text-slate-900">{{ schoolProfile.principalName }}</p>
                            <p class="text-[10px] text-slate-500">NIP. {{ schoolProfile.principalNip }}</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- MODAL: KELOLA KATEGORI -->
        <div v-if="showCategoryModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden flex flex-col">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-tags text-kemensos-600"></i> Kelola Kategori Barang
                    </h3>
                    <button @click="showCategoryModal = false" class="text-slate-400 hover:text-slate-600">
                        <i class="fas fa-xmark"></i>
                    </button>
                </div>
                <div class="p-6 space-y-4 text-xs">
                    <!-- Form Tambah Kategori -->
                    <form @submit.prevent="addCategory" class="flex gap-2">
                        <input v-model="newCategoryName" type="text" placeholder="Nama kategori baru..." required class="flex-1 px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                        <button type="submit" class="px-4 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white font-semibold rounded-xl shadow transition">
                            <i class="fas fa-plus mr-1"></i> Tambah
                        </button>
                    </form>

                    <!-- Daftar Kategori -->
                    <div class="border border-slate-200 rounded-xl divide-y divide-slate-100 max-h-60 overflow-y-auto">
                        <div v-for="(cat, idx) in categories" :key="cat" class="px-3.5 py-2.5 flex items-center justify-between hover:bg-slate-50">
                            <span class="font-medium text-slate-700">{{ cat }}</span>
                            <button @click="removeCategory(idx)" class="text-rose-500 hover:text-rose-700 p-1 rounded-lg hover:bg-rose-50 transition" title="Hapus Kategori">
                                <i class="fas fa-trash-can"></i>
                            </button>
                        </div>
                        <div v-if="categories.length === 0" class="p-4 text-center text-slate-400 italic">
                            Belum ada kategori. Silakan tambahkan kategori di atas.
                        </div>
                    </div>
                </div>
                <div class="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
                    <button @click="showCategoryModal = false" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold rounded-xl">Tutup</button>
                </div>
            </div>
        </div>

        <!-- MODAL: BACKUP & RESTORE DATA -->
        <div v-if="showBackupModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden flex flex-col">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-database text-kemensos-600"></i> Simpan & Backup Data SARPRAS
                    </h3>
                    <button @click="showBackupModal = false" class="text-slate-400 hover:text-slate-600">
                        <i class="fas fa-xmark"></i>
                    </button>
                </div>
                <div class="p-6 space-y-4 text-xs">
                    <!-- Status Simpan LocalStorage -->
                    <div class="p-3.5 bg-emerald-50 border border-emerald-200 rounded-xl flex items-start gap-3">
                        <i class="fas fa-circle-check text-emerald-600 text-lg mt-0.5"></i>
                        <div>
                            <p class="font-bold text-emerald-900">Penyimpanan Otomatis Aktif</p>
                            <p class="text-[11px] text-emerald-700 mt-0.5">Seluruh perubahan data otomatis tersimpan di memori browser perangkat ini.</p>
                        </div>
                    </div>

                    <!-- Export File JSON (Backup Lengkap) -->
                    <div class="p-4 border border-slate-200 rounded-xl space-y-2 bg-slate-50/50">
                        <div class="flex items-center justify-between">
                            <span class="font-bold text-slate-800 flex items-center gap-2">
                                <i class="fas fa-file-code text-blue-600"></i> Download Backup Lengkap (JSON)
                            </span>
                            <span class="text-[10px] bg-blue-100 text-blue-800 font-bold px-2 py-0.5 rounded">Rekomendasi</span>
                        </div>
                        <p class="text-slate-500 text-[11px]">Download seluruh database (Inventaris, Ruangan, Pinjaman, Pemeliharaan, Profil) dalam 1 file backup `.json` untuk disimpan/dipindahkan.</p>
                        <button @click="exportFullJSON" class="w-full mt-1 py-2 px-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl shadow transition flex items-center justify-center gap-2">
                            <i class="fas fa-download"></i> Download File Backup (.json)
                        </button>
                    </div>

                    <!-- Import Backup JSON -->
                    <div class="p-4 border border-slate-200 rounded-xl space-y-2 bg-amber-50/40 border-amber-200">
                        <span class="font-bold text-amber-900 flex items-center gap-2">
                            <i class="fas fa-file-import text-amber-600"></i> Restore / Import File Backup
                        </span>
                        <p class="text-amber-800 text-[11px]">Pulihkan atau masukkan data dari file `.json` backup yang pernah di-download sebelumnya.</p>
                        <input type="file" ref="backupFileInput" @change="importJSON" accept=".json" class="hidden">
                        <button @click="$refs.backupFileInput.click()" class="w-full mt-1 py-2 px-3 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-xl shadow transition flex items-center justify-center gap-2">
                            <i class="fas fa-upload"></i> Pilih File Backup (.json)...
                        </button>
                    </div>
                </div>
                <div class="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
                    <button @click="showBackupModal = false" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold rounded-xl">Tutup</button>
                </div>
            </div>
        </div>

        <!-- MODAL: ADD / EDIT ITEM -->
        <div v-if="showItemModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[90vh]">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between shrink-0">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-box text-kemensos-600"></i>
                        <span>{{ isEditing ? 'Edit Data Barang' : 'Tambah Barang Inventaris Baru' }}</span>
                    </h3>
                    <button @click="showItemModal = false" class="text-slate-400 hover:text-slate-600 rounded-lg p-1 transition">
                        <i class="fas fa-xmark text-base"></i>
                    </button>
                </div>

                <form @submit.prevent="saveItem" class="p-6 space-y-4 overflow-y-auto text-xs">
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Kode Barang *</label>
                            <input v-model="formItem.code" type="text" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none font-mono">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Nama Barang *</label>
                            <input v-model="formItem.name" type="text" required placeholder="Contoh: Laptop Chromebook" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Kategori *</label>
                            <select v-model="formItem.category" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                                <option v-for="c in categories" :key="c" :value="c">{{ c }}</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Lokasi Ruangan *</label>
                            <select v-model="formItem.room" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                                <option v-for="r in rooms" :key="r.id" :value="r.name">{{ r.name }}</option>
                            </select>
                        </div>
                    </div>

                    <div class="grid grid-cols-3 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Jumlah Unit *</label>
                            <input v-model.number="formItem.qty" type="number" min="1" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Satuan *</label>
                            <input v-model="formItem.unit" type="text" required placeholder="Unit, Set, Pcs" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Kondisi *</label>
                            <select v-model="formItem.condition" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                                <option value="Baik">Baik</option>
                                <option value="Rusak Ringan">Rusak Ringan</option>
                                <option value="Rusak Berat">Rusak Berat</option>
                            </select>
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Sumber Anggaran *</label>
                            <select v-model="formItem.fundSource" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                                <option v-for="f in fundSources" :key="f" :value="f">{{ f }}</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Tahun Pengadaan *</label>
                            <input v-model.number="formItem.year" type="number" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                        </div>
                    </div>

                    <div class="pt-4 border-t border-slate-100 flex items-center justify-end gap-3">
                        <button type="button" @click="showItemModal = false" class="px-4 py-2 rounded-xl text-slate-600 hover:bg-slate-100 font-semibold">Batal</button>
                        <button type="submit" class="px-5 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white rounded-xl font-semibold shadow transition">Simpan Barang</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: BORROW -->
        <div v-if="showBorrowModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[90vh]">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between shrink-0">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-hand-holding-hand text-kemensos-600"></i> Pinjam Pakai Peralatan
                    </h3>
                    <button @click="showBorrowModal = false" class="text-slate-400 hover:text-slate-600">
                        <i class="fas fa-xmark"></i>
                    </button>
                </div>
                <form @submit.prevent="saveBorrowing" class="p-6 space-y-4 overflow-y-auto text-xs">
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Nama Peminjam *</label>
                        <input v-model="formBorrow.borrowerName" type="text" required placeholder="Nama Siswa / Pendamping / Tutor" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none">
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Peran / Jabatan *</label>
                            <select v-model="formBorrow.role" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                                <option value="Pendamping Sosial">Pendamping Sosial</option>
                                <option value="Tutor / Pengajar">Tutor / Pengajar</option>
                                <option value="Siswa / Anak Asuh">Siswa / Anak Asuh</option>
                                <option value="Pengelola Ruangan">Pengelola Ruangan</option>
                            </select>
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Tanggal Pinjam *</label>
                            <input v-model="formBorrow.borrowDate" type="date" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                        </div>
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Rencana Pengembalian *</label>
                        <input v-model="formBorrow.returnPlanDate" type="date" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                    </div>

                    <!-- MULTI ITEM SELECTION SECTION -->
                    <div class="border-t border-slate-200 pt-3 space-y-3">
                        <div class="flex items-center justify-between">
                            <label class="font-bold text-slate-800 flex items-center gap-1.5">
                                <i class="fas fa-boxes-stacked text-kemensos-600"></i> Daftar Barang Dipinjam *
                            </label>
                            <button type="button" @click="addBorrowItemRow" class="text-kemensos-600 hover:text-kemensos-700 font-bold text-[11px] flex items-center gap-1 bg-rose-50 hover:bg-rose-100 px-2.5 py-1 rounded-lg border border-rose-200 transition">
                                <i class="fas fa-plus"></i> Tambah Barang
                            </button>
                        </div>

                        <div v-for="(itemRow, idx) in formBorrow.items" :key="idx" class="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2 relative">
                            <div class="flex items-center gap-2">
                                <div class="flex-1">
                                    <label class="block text-[10px] font-semibold text-slate-500 mb-0.5">Barang {{ idx + 1 }}</label>
                                    <select v-model="itemRow.itemId" required class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-kemensos-500 bg-white">
                                        <option v-for="i in availableItemsForBorrow" :key="i.id" :value="i.id">
                                            {{ i.name }} (Stok: {{ i.qty }} {{ i.unit }})
                                        </option>
                                    </select>
                                </div>
                                <div class="w-24">
                                    <label class="block text-[10px] font-semibold text-slate-500 mb-0.5">Jumlah</label>
                                    <input v-model.number="itemRow.qty" type="number" min="1" required class="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-kemensos-500 bg-white">
                                </div>
                                <button v-if="formBorrow.items.length > 1" type="button" @click="removeBorrowItemRow(idx)" class="mt-4 p-1.5 text-rose-600 hover:bg-rose-100 rounded-lg transition" title="Hapus Baris">
                                    <i class="fas fa-trash-can"></i>
                                </button>
                            </div>
                        </div>
                    </div>

                    <div class="pt-4 border-t border-slate-100 flex items-center justify-end gap-3">
                        <button type="button" @click="showBorrowModal = false" class="px-4 py-2 rounded-xl text-slate-600 hover:bg-slate-100">Batal</button>
                        <button type="submit" class="px-5 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white rounded-xl font-semibold shadow">Simpan Peminjaman</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: SCHOOL PROFILE SETTINGS -->
        <div v-if="showSchoolModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden flex flex-col max-h-[90vh]">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-landmark text-kemensos-600"></i> Identitas Sekolah Rakyat Kemensos
                    </h3>
                    <button @click="showSchoolModal = false" class="text-slate-400 hover:text-slate-600">
                        <i class="fas fa-xmark"></i>
                    </button>
                </div>
                <form @submit.prevent="saveSchoolProfile" class="p-6 space-y-4 overflow-y-auto text-xs">
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Nama Unit / Sekolah Rakyat *</label>
                        <input v-model="formSchool.name" type="text" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                    </div>

                    <!-- Input Alamat Sekolah -->
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Alamat Lengkap Sekolah / Sentra *</label>
                        <textarea v-model="formSchool.address" rows="2" required placeholder="Contoh: Jl. Dewi Sartika No. 12, Cawang, Jakarta Selatan" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500 focus:outline-none"></textarea>
                    </div>

                    <!-- Upload Logo Section -->
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Logo Sekolah / Sentra</label>
                        <div class="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-xl">
                            <div class="w-12 h-12 rounded-xl bg-white flex items-center justify-center overflow-hidden shrink-0 border border-slate-200 shadow-sm">
                                <img v-if="formSchool.logoUrl" :src="formSchool.logoUrl" class="w-full h-full object-cover">
                                <i v-else class="fas fa-hands-holding-child text-xl text-kemensos-700"></i>
                            </div>
                            <div class="flex-1 min-w-0">
                                <input type="file" @change="onLogoSelected" accept="image/*" class="w-full text-[11px] text-slate-600 file:mr-2 file:py-1 file:px-2.5 file:rounded-lg file:border-0 file:text-[11px] file:font-semibold file:bg-kemensos-600 file:text-white hover:file:bg-kemensos-700 cursor-pointer">
                                <p class="text-[10px] text-slate-400 mt-1">Format: PNG, JPG, JPEG (Maksimal 2MB)</p>
                            </div>
                            <button v-if="formSchool.logoUrl" type="button" @click="removeLogo" class="text-xs text-rose-600 hover:text-rose-700 font-bold px-2.5 py-1.5 bg-rose-50 hover:bg-rose-100 rounded-lg border border-rose-200 transition shrink-0">
                                Hapus
                            </button>
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Kota / Kabupaten *</label>
                            <input v-model="formSchool.city" type="text" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Email</label>
                            <input v-model="formSchool.email" type="email" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                        </div>
                    </div>
                    <div class="border-t border-slate-100 pt-3 font-bold text-slate-800">Pejabat / Penandatangan:</div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Nama Kepala Sentra/Sekolah</label>
                            <input v-model="formSchool.principalName" type="text" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">NIP Kepala</label>
                            <input v-model="formSchool.principalNip" type="text" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Nama Pengelola SARPRAS</label>
                            <input v-model="formSchool.sarprasHeadName" type="text" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                        </div>
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">NIP Pengelola</label>
                            <input v-model="formSchool.sarprasHeadNip" type="text" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                        </div>
                    </div>
                    <div class="pt-4 border-t border-slate-100 flex items-center justify-end gap-3">
                        <button type="button" @click="showSchoolModal = false" class="px-4 py-2 rounded-xl text-slate-600 hover:bg-slate-100">Batal</button>
                        <button type="submit" class="px-5 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white rounded-xl font-semibold shadow">Simpan Identitas</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: USER PROFILE SETTINGS -->
        <div v-if="showAccountModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden flex flex-col">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h3 class="font-bold text-slate-800 text-sm flex items-center gap-2">
                        <i class="fas fa-user-gear text-kemensos-600"></i> Edit Profil Operator
                    </h3>
                    <button @click="showAccountModal = false" class="text-slate-400 hover:text-slate-600">
                        <i class="fas fa-xmark"></i>
                    </button>
                </div>
                <form @submit.prevent="saveUserProfile" class="p-6 space-y-4 text-xs">
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Nama Operator / Petugas *</label>
                        <input v-model="formUser.name" type="text" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Peran / Jabatan *</label>
                        <input v-model="formUser.role" type="text" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Email</label>
                        <input v-model="formUser.email" type="email" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                    </div>
                    <div class="pt-4 border-t border-slate-100 flex items-center justify-end gap-3">
                        <button type="button" @click="showAccountModal = false" class="px-4 py-2 rounded-xl text-slate-600 hover:bg-slate-100">Batal</button>
                        <button type="submit" class="px-5 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white rounded-xl font-semibold shadow">Simpan Profil</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: CONFIRM CLEAR ALL DATA -->
        <div v-if="showClearModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-sm w-full p-6 text-center">
                <div class="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center text-xl mx-auto mb-3">
                    <i class="fas fa-triangle-exclamation"></i>
                </div>
                <h3 class="font-bold text-slate-900 text-sm">Kosongkan Semua Data?</h3>
                <p class="text-xs text-slate-500 mt-1">Seluruh data inventaris, pinjaman, perbaikan, dan ruangan akan dihapus secara permanen. Tindakan ini tidak dapat dibatalkan.</p>
                <div class="mt-5 flex gap-3 justify-center">
                    <button @click="showClearModal = false" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl">Batal</button>
                    <button @click="executeClearAllData" class="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold rounded-xl shadow">Ya, Kosongkan Data</button>
                </div>
            </div>
        </div>

        <!-- MODAL: ADD / EDIT ROOM -->
        <div v-if="showRoomModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden flex flex-col">
                <div class="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h3 class="font-bold text-slate-800 text-sm">{{ isEditingRoom ? 'Edit Data Ruangan' : 'Tambah Ruangan Baru' }}</h3>
                    <button @click="showRoomModal = false" class="text-slate-400 hover:text-slate-600">
                        <i class="fas fa-xmark"></i>
                    </button>
                </div>
                <form @submit.prevent="saveRoom" class="p-6 space-y-4 text-xs">
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Nama Ruangan *</label>
                        <input v-model="formRoom.name" type="text" required placeholder="Contoh: Bengkel Otomotif Vokasi" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                    </div>
                    <div>
                        <label class="block font-bold text-slate-700 mb-1">Penanggung Jawab (PIC)</label>
                        <input v-model="formRoom.pic" type="text" placeholder="Nama Petugas / Pendamping" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-kemensos-500">
                    </div>
                    <div class="pt-4 border-t border-slate-100 flex items-center justify-end gap-3">
                        <button type="button" @click="showRoomModal = false" class="px-4 py-2 rounded-xl text-slate-600 hover:bg-slate-100">Batal</button>
                        <button type="submit" class="px-5 py-2 bg-kemensos-600 hover:bg-kemensos-700 text-white rounded-xl font-semibold shadow">Simpan Ruangan</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL: CONFIRM DELETE ROOM -->
        <div v-if="showDeleteRoomModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-sm w-full p-6 text-center">
                <div class="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center text-xl mx-auto mb-3">
                    <i class="fas fa-trash"></i>
                </div>
                <h3 class="font-bold text-slate-900 text-sm">Hapus Ruangan Ini?</h3>
                <p class="text-xs text-slate-500 mt-1">Tindakan ini tidak dapat dibatalkan.</p>
                <div class="mt-5 flex gap-3 justify-center">
                    <button @click="showDeleteRoomModal = false" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl">Batal</button>
                    <button @click="executeDeleteRoom" class="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold rounded-xl shadow">Hapus</button>
                </div>
            </div>
        </div>

        <!-- MODAL: CONFIRM DELETE SINGLE ITEM -->
        <div v-if="showDeleteModal" class="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm no-print">
            <div class="bg-white rounded-2xl shadow-2xl max-w-sm w-full p-6 text-center">
                <div class="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center text-xl mx-auto mb-3">
                    <i class="fas fa-trash"></i>
                </div>
                <h3 class="font-bold text-slate-900 text-sm">Hapus Barang Ini?</h3>
                <p class="text-xs text-slate-500 mt-1">Tindakan ini tidak dapat dibatalkan.</p>
                <div class="mt-5 flex gap-3 justify-center">
                    <button @click="showDeleteModal = false" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl">Batal</button>
                    <button @click="executeDeleteItem" class="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold rounded-xl shadow">Hapus</button>
                </div>
            </div>
        </div>

        <!-- TOAST NOTIFICATION -->
        <div v-if="toast.show" :class="toast.type === 'error' ? 'bg-rose-600' : 'bg-slate-900'" class="fixed bottom-5 right-5 z-50 text-white px-4 py-3 rounded-2xl shadow-2xl text-xs font-semibold flex items-center gap-2 transition transform animate-bounce">
            <i :class="toast.type === 'error' ? 'fa-circle-xmark' : 'fa-circle-check'"></i>
            <span>{{ toast.message }}</span>
        </div>

    </div>

    <!-- VUE 3 LOGIC & SCRIPT -->
    <script>
        const { createApp } = Vue;

        const defaultSchoolProfile = {
            name: 'Sekolah Rakyat Kemensos Sentra Jakarta',
            satker: 'Direktorat Jenderal Rehabilitasi Sosial KEMENSOS RI',
            city: 'Jakarta Selatan',
            address: 'Jl. Dewi Sartika No. 12, Cawang, Jakarta Selatan',
            email: 'sekolahrakyat@kemensos.go.id',
            principalName: 'Drs. H. M. Yasin, M.Si.',
            principalNip: '19720315 199803 1 003',
            sarprasHeadName: 'Rahmat Hidayat, S.Sos.',
            sarprasHeadNip: '19850620 201012 1 002',
            logoUrl: null
        };

        const defaultUserProfile = {
            name: 'Rahmat Hidayat',
            role: 'Pengelola SARPRAS Kemensos',
            email: 'rahmat.sarpras@kemensos.go.id',
            avatarText: 'RH'
        };

        // Empty default arrays
        const sampleInventory = [];
        const sampleRooms = [];
        const sampleBorrowings = [];
        const sampleMaintenances = [];

        const defaultCategories = [
            'Perlengkapan Belajar & TIK',
            'Alat Vokasional / Keterampilan',
            'Mebel & Furnitur',
            'Alat Olahraga & Seni',
            'Peralatan Dapur / Konsumsi',
            'Sanitasi & Kesehatan',
            'Lainnya'
        ];

        function loadStorage(key, fallback) {
            try {
                const item = localStorage.getItem(key);
                return item ? JSON.parse(item) : fallback;
            } catch (e) {
                return fallback;
            }
        }

        createApp({
            data() {
                return {
                    currentTab: 'dashboard',
                    showUserDropdown: false,

                    schoolProfile: loadStorage('sarpras_sr_school_profile', defaultSchoolProfile),
                    userProfile: loadStorage('sarpras_sr_user_profile', defaultUserProfile),

                    inventory: loadStorage('sarpras_sr_inventory', sampleInventory),
                    rooms: loadStorage('sarpras_sr_rooms', sampleRooms),
                    borrowings: loadStorage('sarpras_sr_borrowings', sampleBorrowings),
                    maintenances: loadStorage('sarpras_sr_maintenances', sampleMaintenances),
                    categories: loadStorage('sarpras_sr_categories', defaultCategories),

                    fundSources: [
                        'DIPA Kemensos',
                        'Bantuan Hibah Sosial',
                        'APBN Kemensos',
                        'Sumbangan Donatur / LSM'
                    ],

                    invSearch: '',
                    invCategoryFilter: '',
                    invConditionFilter: '',

                    showItemModal: false,
                    showBorrowModal: false,
                    showMaintenanceModal: false,
                    showRoomModal: false,
                    showCategoryModal: false,
                    showSchoolModal: false,
                    showAccountModal: false,
                    showBackupModal: false,
                    showClearModal: false,
                    showDeleteModal: false,
                    showDeleteRoomModal: false,
                    showBorrowLetterModal: false,
                    showImportExcelModal: false,
                    
                    newCategoryName: '',
                    itemToDeleteId: null,
                    itemToDelete: null,
                    roomToDeleteId: null,
                    roomToDelete: null,
                    selectedBorrowing: null,
                    isEditing: false,
                    isEditingRoom: false,

                    parsedExcelItems: [],
                    excelImportMode: 'append',

                    formSchool: { ...defaultSchoolProfile },
                    formUser: { ...defaultUserProfile },
                    formItem: { id: null, code: '', name: '', category: '', room: '', qty: 1, unit: 'Unit', condition: 'Baik', fundSource: 'DIPA Kemensos', year: new Date().getFullYear(), notes: '' },
                    formBorrow: { borrowerName: '', role: 'Siswa / Anak Asuh', borrowDate: new Date().toISOString().split('T')[0], returnPlanDate: '', items: [{ itemId: '', qty: 1 }] },
                    formMaintenance: { itemId: '', reporter: '', description: '', cost: 0, targetCondition: 'Rusak Ringan' },
                    formRoom: { id: null, name: '', pic: '' },

                    toast: { show: false, message: '', type: 'success' },

                    conditionChartInstance: null,
                    categoryChartInstance: null
                };
            },

            computed: {
                filteredInventory() {
                    return this.inventory.filter(i => {
                        const searchLower = (this.invSearch || '').toLowerCase();
                        const matchesSearch = (i.name || '').toLowerCase().includes(searchLower) || (i.code || '').toLowerCase().includes(searchLower);
                        const matchesCat = !this.invCategoryFilter || i.category === this.invCategoryFilter;
                        const matchesCond = !this.invConditionFilter || i.condition === this.invConditionFilter;
                        return matchesSearch && matchesCat && matchesCond;
                    });
                },

                availableItemsForBorrow() {
                    return this.inventory.filter(i => Number(i.qty) > 0 && i.condition === 'Baik');
                },

                totalItemsCount() {
                    return this.inventory.reduce((sum, i) => sum + Number(i.qty || 0), 0);
                },

                goodConditionCount() {
                    return this.inventory.filter(i => i.condition === 'Baik').reduce((sum, i) => sum + Number(i.qty || 0), 0);
                },

                damagedCount() {
                    return this.inventory.filter(i => i.condition !== 'Baik').reduce((sum, i) => sum + Number(i.qty || 0), 0);
                },

                activeBorrowingsCount() {
                    return this.borrowings.filter(b => b.status === 'Dipinjam').length;
                },

                roomBreakdown() {
                    return this.rooms.map(r => {
                        const roomItems = this.inventory.filter(i => i.room === r.name);
                        const itemTypesCount = roomItems.length;
                        const totalUnits = roomItems.reduce((sum, i) => sum + Number(i.qty || 0), 0);
                        return { ...r, itemTypesCount, totalUnits };
                    });
                }
            },

            watch: {
                inventory: { handler() { localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory)); this.updateCharts(); }, deep: true },
                rooms: { handler() { localStorage.setItem('sarpras_sr_rooms', JSON.stringify(this.rooms)); }, deep: true },
                borrowings: { handler() { localStorage.setItem('sarpras_sr_borrowings', JSON.stringify(this.borrowings)); }, deep: true },
                maintenances: { handler() { localStorage.setItem('sarpras_sr_maintenances', JSON.stringify(this.maintenances)); }, deep: true },
                categories: { handler() { localStorage.setItem('sarpras_sr_categories', JSON.stringify(this.categories)); }, deep: true },
                schoolProfile: { handler() { localStorage.setItem('sarpras_sr_school_profile', JSON.stringify(this.schoolProfile)); }, deep: true },
                userProfile: { handler() { localStorage.setItem('sarpras_sr_user_profile', JSON.stringify(this.userProfile)); }, deep: true },
                currentTab(newTab) {
                    if (newTab === 'dashboard') {
                        this.$nextTick(() => this.renderCharts());
                    }
                }
            },

            mounted() {
                // Ensure all loaded records have an explicit ID
                this.inventory.forEach((i, idx) => { if (!i.id) i.id = Date.now() + idx; });
                this.rooms.forEach((r, idx) => { if (!r.id) r.id = Date.now() + idx; });
                this.borrowings.forEach((b, idx) => { if (!b.id) b.id = Date.now() + idx; });
                this.maintenances.forEach((m, idx) => { if (!m.id) m.id = Date.now() + idx; });

                this.$nextTick(() => {
                    this.renderCharts();
                });
            },

            methods: {
                formatCurrency(val) {
                    return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(val || 0);
                },

                formatDate(dateStr) {
                    if (!dateStr) return '-';
                    const d = new Date(dateStr);
                    if (isNaN(d.getTime())) return dateStr;
                    return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
                },

                getItemName(id) {
                    const item = this.inventory.find(i => i.id === id);
                    return item ? item.name : 'Barang Dihapus';
                },

                getItemObject(id) {
                    return this.inventory.find(i => i.id === id) || { code: '-', name: 'Barang Dihapus', category: '-', room: '-', unit: 'Unit' };
                },

                getBorrowingItemsList(b) {
                    if (!b) return [];
                    if (b.items && Array.isArray(b.items)) {
                        return b.items.map(row => {
                            const itemObj = this.getItemObject(row.itemId);
                            return {
                                ...itemObj,
                                qtyBorrowed: row.qty
                            };
                        });
                    } else if (b.itemId) {
                        const itemObj = this.getItemObject(b.itemId);
                        return [{
                            ...itemObj,
                            qtyBorrowed: b.qty || 1
                        }];
                    }
                    return [];
                },

                openBorrowLetter(b) {
                    this.selectedBorrowing = b;
                    this.showBorrowLetterModal = true;
                },

                openImportExcelModal() {
                    this.parsedExcelItems = [];
                    if (this.$refs.excelFileInput) {
                        this.$refs.excelFileInput.value = '';
                    }
                    this.showImportExcelModal = true;
                },

                downloadExcelTemplate() {
                    const headers = ['Kode Barang', 'Nama Barang', 'Kategori', 'Ruangan', 'Jumlah', 'Satuan', 'Kondisi', 'Sumber Anggaran', 'Tahun Pengadaan'];
                    const sampleRows = [
                        ['SRMP-INV-001', 'Laptop Chromebook Belajar', 'Perlengkapan Belajar & TIK', 'Ruang Belajar TIK', 10, 'Unit', 'Baik', 'DIPA Kemensos', 2024],
                        ['SRMP-INV-002', 'Mesin Jahit Konveksi', 'Alat Vokasional / Keterampilan', 'Bengkel Vokasi', 5, 'Unit', 'Baik', 'Bantuan Hibah Sosial', 2024],
                        ['SRMP-INV-003', 'Meja Kursi Siswa', 'Mebel & Furnitur', 'Ruang Kelas 1', 20, 'Set', 'Baik', 'DIPA Kemensos', 2023]
                    ];
                    
                    const data = [headers, ...sampleRows];
                    const ws = XLSX.utils.aoa_to_sheet(data);
                    
                    ws['!cols'] = [
                        { wch: 16 },
                        { wch: 32 },
                        { wch: 28 },
                        { wch: 22 },
                        { wch: 10 },
                        { wch: 10 },
                        { wch: 12 },
                        { wch: 22 },
                        { wch: 16 }
                    ];

                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, 'Template SARPRAS');
                    XLSX.writeFile(wb, 'Template_Import_SARPRAS_Sekolah_Rakyat.xlsx');
                    this.showToast('Template Excel (.xlsx) berhasil di-download!');
                },

                handleExcelFileUpload(event) {
                    const file = event.target.files[0];
                    if (!file) return;

                    const reader = new FileReader();
                    reader.onload = (e) => {
                        try {
                            const data = e.target.result;
                            const workbook = XLSX.read(data, { type: 'binary' });
                            const firstSheetName = workbook.SheetNames[0];
                            const worksheet = workbook.Sheets[firstSheetName];
                            const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

                            if (jsonRows.length === 0) {
                                this.showToast('File Excel kosong atau format tidak sesuai', 'error');
                                return;
                            }

                            this.parsedExcelItems = jsonRows.map((row, idx) => {
                                const code = row['Kode Barang'] || row['Kode'] || row['kode'] || `SRMP-INV-${String(this.inventory.length + idx + 1).padStart(3, '0')}`;
                                const name = row['Nama Barang'] || row['Nama'] || row['nama'] || 'Barang Tanpa Nama';
                                const category = row['Kategori'] || row['kategori'] || (this.categories[0] || 'Lainnya');
                                const room = row['Ruangan'] || row['Lokasi'] || row['ruangan'] || (this.rooms[0] ? this.rooms[0].name : 'Ruang Belajar');
                                const qty = Number(row['Jumlah'] || row['Qty'] || row['qty'] || 1);
                                const unit = row['Satuan'] || row['satuan'] || 'Unit';
                                const condition = row['Kondisi'] || row['kondisi'] || 'Baik';
                                const fundSource = row['Sumber Anggaran'] || row['Sumber'] || 'DIPA Kemensos';
                                const year = Number(row['Tahun Pengadaan'] || row['Tahun'] || new Date().getFullYear());

                                return {
                                    id: Date.now() + idx,
                                    code: String(code).trim(),
                                    name: String(name).trim(),
                                    category: String(category).trim(),
                                    room: String(room).trim(),
                                    qty: isNaN(qty) ? 1 : qty,
                                    unit: String(unit).trim(),
                                    condition: String(condition).trim(),
                                    fundSource: String(fundSource).trim(),
                                    year: isNaN(year) ? new Date().getFullYear() : year
                                };
                            });

                            this.showToast(`${this.parsedExcelItems.length} baris data Excel berhasil dibaca!`);
                        } catch (err) {
                            this.showToast('Gagal membaca file Excel: ' + err.message, 'error');
                        }
                    };
                    reader.readAsBinaryString(file);
                },

                executeImportExcel() {
                    if (this.parsedExcelItems.length === 0) {
                        this.showToast('Tidak ada data Excel yang siap di-import', 'error');
                        return;
                    }

                    if (this.excelImportMode === 'replace') {
                        this.inventory = [...this.parsedExcelItems];
                    } else {
                        this.inventory = [...this.parsedExcelItems, ...this.inventory];
                    }

                    // Auto register new categories or rooms if not present
                    this.parsedExcelItems.forEach(item => {
                        if (item.category && !this.categories.includes(item.category)) {
                            this.categories.push(item.category);
                        }
                        if (item.room && !this.rooms.some(r => r.name === item.room)) {
                            this.rooms.push({ id: Date.now() + Math.random(), name: item.room, pic: 'Import Excel' });
                        }
                    });

                    localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory));
                    localStorage.setItem('sarpras_sr_categories', JSON.stringify(this.categories));
                    localStorage.setItem('sarpras_sr_rooms', JSON.stringify(this.rooms));

                    this.showToast(`Berhasil meng-import ${this.parsedExcelItems.length} data barang dari Excel!`);
                    this.showImportExcelModal = false;
                    this.parsedExcelItems = [];
                    this.$nextTick(() => this.updateCharts());
                },

                printBorrowLetter() {
                    window.print();
                },

                showToast(message, type = 'success') {
                    this.toast = { show: true, message, type };
                    setTimeout(() => { this.toast.show = false; }, 3000);
                },

                openAddModal() {
                    this.isEditing = false;
                    const nextNum = (this.inventory.length + 1).toString().padStart(3, '0');
                    this.formItem = {
                        id: null,
                        code: `SRMP-INV-${nextNum}`,
                        name: '',
                        category: this.categories[0] || 'Lainnya',
                        room: this.rooms[0] ? this.rooms[0].name : 'Ruang Belajar',
                        qty: 1,
                        unit: 'Unit',
                        condition: 'Baik',
                        fundSource: this.fundSources[0],
                        year: new Date().getFullYear(),
                        notes: ''
                    };
                    this.showItemModal = true;
                },

                openCategoryModal() {
                    this.newCategoryName = '';
                    this.showCategoryModal = true;
                },

                addCategory() {
                    const name = (this.newCategoryName || '').trim();
                    if (!name) return;
                    if (this.categories.includes(name)) {
                        this.showToast('Kategori ini sudah ada!', 'error');
                        return;
                    }
                    this.categories.push(name);
                    this.newCategoryName = '';
                    this.showToast('Kategori baru berhasil ditambahkan');
                },

                removeCategory(index) {
                    const removed = this.categories[index];
                    this.categories.splice(index, 1);
                    this.showToast(`Kategori "${removed}" telah dihapus`, 'error');
                },

                editItem(item) {
                    this.isEditing = true;
                    this.formItem = JSON.parse(JSON.stringify(item));
                    if (!this.formItem.id) {
                        this.formItem.id = item.id || Date.now();
                        item.id = this.formItem.id;
                    }
                    this.showItemModal = true;
                },

                saveItem() {
                    if (!this.formItem.name || !this.formItem.code) {
                        this.showToast('Mohon lengkapi kode dan nama barang', 'error');
                        return;
                    }

                    if (this.isEditing) {
                        const targetId = String(this.formItem.id || '');
                        const targetCode = String(this.formItem.code || '');
                        
                        let idx = -1;
                        if (targetId) {
                            idx = this.inventory.findIndex(i => String(i.id || '') === targetId);
                        }
                        if (idx === -1 && targetCode) {
                            idx = this.inventory.findIndex(i => String(i.code || '') === targetCode);
                        }

                        if (idx !== -1) {
                            this.inventory.splice(idx, 1, { ...this.formItem });
                            this.showToast('Data barang berhasil diperbarui');
                        } else {
                            const newItem = { ...this.formItem, id: this.formItem.id || Date.now() };
                            this.inventory.unshift(newItem);
                            this.showToast('Data barang diperbarui');
                        }
                    } else {
                        const newItem = { ...this.formItem, id: Date.now() };
                        this.inventory.unshift(newItem);
                        this.showToast('Barang baru berhasil ditambahkan');
                    }

                    this.inventory = [...this.inventory];
                    localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory));
                    this.showItemModal = false;
                    this.isEditing = false;
                    this.$nextTick(() => this.updateCharts());
                },

                confirmDeleteItem(item) {
                    if (typeof item === 'object' && item !== null) {
                        this.itemToDelete = item;
                        this.itemToDeleteId = item.id || item.code;
                    } else {
                        this.itemToDelete = null;
                        this.itemToDeleteId = item;
                    }
                    this.showDeleteModal = true;
                },

                executeDeleteItem() {
                    if (this.itemToDelete) {
                        const idx = this.inventory.indexOf(this.itemToDelete);
                        if (idx !== -1) {
                            this.inventory.splice(idx, 1);
                        } else {
                            const targetId = String(this.itemToDeleteId || '');
                            this.inventory = this.inventory.filter(i => String(i.id || '') !== targetId && String(i.code || '') !== targetId);
                        }
                    } else if (this.itemToDeleteId !== null && this.itemToDeleteId !== undefined) {
                        const targetId = String(this.itemToDeleteId);
                        this.inventory = this.inventory.filter(i => String(i.id || '') !== targetId && String(i.code || '') !== targetId);
                    }

                    this.inventory = [...this.inventory];
                    localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory));
                    this.showToast('Barang berhasil dihapus', 'error');

                    this.showDeleteModal = false;
                    this.itemToDeleteId = null;
                    this.itemToDelete = null;
                    this.$nextTick(() => this.updateCharts());
                },

                openBorrowModal() {
                    if (this.availableItemsForBorrow.length === 0) {
                        this.showToast('Tidak ada barang tersedia untuk dipinjam', 'error');
                        return;
                    }
                    const today = new Date();
                    const defaultReturn = new Date();
                    defaultReturn.setDate(today.getDate() + 7);

                    this.formBorrow = {
                        borrowerName: '',
                        role: 'Siswa / Anak Asuh',
                        borrowDate: today.toISOString().split('T')[0],
                        returnPlanDate: defaultReturn.toISOString().split('T')[0],
                        items: [{ itemId: this.availableItemsForBorrow[0].id, qty: 1 }]
                    };
                    this.showBorrowModal = true;
                },

                addBorrowItemRow() {
                    const defaultId = this.availableItemsForBorrow.length > 0 ? this.availableItemsForBorrow[0].id : '';
                    this.formBorrow.items.push({ itemId: defaultId, qty: 1 });
                },

                removeBorrowItemRow(idx) {
                    if (this.formBorrow.items.length > 1) {
                        this.formBorrow.items.splice(idx, 1);
                    }
                },

                saveBorrowing() {
                    if (!this.formBorrow.borrowerName) {
                        this.showToast('Mohon isi nama peminjam', 'error');
                        return;
                    }
                    if (!this.formBorrow.items || this.formBorrow.items.length === 0) {
                        this.showToast('Pilih setidaknya 1 barang untuk dipinjam', 'error');
                        return;
                    }

                    // Validate stock for each item row
                    for (const row of this.formBorrow.items) {
                        if (!row.itemId) {
                            this.showToast('Pilih barang yang ingin dipinjam', 'error');
                            return;
                        }
                        const targetItem = this.inventory.find(i => i.id === row.itemId);
                        if (!targetItem) {
                            this.showToast('Barang tidak ditemukan', 'error');
                            return;
                        }
                        if (Number(row.qty) <= 0) {
                            this.showToast('Jumlah pinjam minimal 1', 'error');
                            return;
                        }
                        if (Number(row.qty) > Number(targetItem.qty)) {
                            this.showToast(`Stok "${targetItem.name}" tidak mencukupi (Tersedia: ${targetItem.qty})`, 'error');
                            return;
                        }
                    }

                    // Deduct stock for each borrowed item
                    for (const row of this.formBorrow.items) {
                        const targetItem = this.inventory.find(i => i.id === row.itemId);
                        if (targetItem) {
                            targetItem.qty = Number(targetItem.qty) - Number(row.qty);
                        }
                    }

                    this.borrowings.unshift({
                        id: Date.now(),
                        borrowerName: this.formBorrow.borrowerName,
                        role: this.formBorrow.role,
                        borrowDate: this.formBorrow.borrowDate,
                        returnPlanDate: this.formBorrow.returnPlanDate,
                        items: JSON.parse(JSON.stringify(this.formBorrow.items)),
                        status: 'Dipinjam'
                    });

                    this.inventory = [...this.inventory];
                    localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory));
                    this.showToast('Peminjaman barang berhasil dicatat');
                    this.showBorrowModal = false;
                },

                returnItem(b) {
                    if (b.items && Array.isArray(b.items)) {
                        b.items.forEach(row => {
                            const targetItem = this.inventory.find(i => i.id === row.itemId);
                            if (targetItem) {
                                targetItem.qty = Number(targetItem.qty) + Number(row.qty);
                            }
                        });
                    } else if (b.itemId) {
                        const targetItem = this.inventory.find(i => i.id === b.itemId);
                        if (targetItem) {
                            targetItem.qty = Number(targetItem.qty) + Number(b.qty);
                        }
                    }
                    b.status = 'Dikembalikan';
                    this.inventory = [...this.inventory];
                    localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory));
                    this.showToast('Seluruh barang pinjaman telah dikembalikan');
                },

                openMaintenanceModal() {
                    if (this.inventory.length === 0) {
                        this.showToast('Belum ada data barang', 'error');
                        return;
                    }
                    this.formMaintenance = {
                        itemId: this.inventory[0].id,
                        reporter: '',
                        description: '',
                        cost: 0,
                        targetCondition: 'Rusak Ringan'
                    };
                    this.showMaintenanceModal = true;
                },

                saveMaintenance() {
                    const targetItem = this.inventory.find(i => i.id === this.formMaintenance.itemId);
                    if (targetItem) {
                        targetItem.condition = this.formMaintenance.targetCondition;
                    }

                    this.maintenances.unshift({
                        id: Date.now(),
                        ...this.formMaintenance,
                        date: new Date().toISOString().split('T')[0],
                        status: 'Menunggu'
                    });

                    this.showToast('Laporan kerusakan tersimpan');
                    this.showMaintenanceModal = false;
                },

                updateMaintenanceStatus(m, newStatus) {
                    m.status = newStatus;
                    if (newStatus === 'Selesai') {
                        const targetItem = this.inventory.find(i => i.id === m.itemId);
                        if (targetItem) targetItem.condition = 'Baik';
                        this.showToast('Perbaikan selesai. Kondisi barang kini: Baik');
                    }
                },

                openRoomModal() {
                    this.isEditingRoom = false;
                    this.formRoom = { id: null, name: '', pic: '' };
                    this.showRoomModal = true;
                },

                editRoom(r) {
                    this.isEditingRoom = true;
                    this.formRoom = JSON.parse(JSON.stringify(r));
                    if (!this.formRoom.id) {
                        this.formRoom.id = r.id || Date.now();
                        r.id = this.formRoom.id;
                    }
                    this.showRoomModal = true;
                },

                saveRoom() {
                    if (!this.formRoom.name) return;

                    if (this.isEditingRoom) {
                        const targetId = String(this.formRoom.id || '');
                        let idx = -1;
                        if (targetId) {
                            idx = this.rooms.findIndex(r => String(r.id || '') === targetId);
                        }
                        if (idx === -1) {
                            idx = this.rooms.findIndex(r => r.name === this.formRoom.name);
                        }

                        if (idx !== -1) {
                            const oldName = this.rooms[idx].name;
                            const newName = this.formRoom.name;
                            if (oldName !== newName) {
                                this.inventory.forEach(item => {
                                    if (item.room === oldName) item.room = newName;
                                });
                                localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory));
                            }
                            this.rooms.splice(idx, 1, { ...this.formRoom });
                            this.showToast('Data ruangan diperbarui');
                        } else {
                            this.rooms.push({ ...this.formRoom, id: this.formRoom.id || Date.now() });
                            this.showToast('Data ruangan diperbarui');
                        }
                    } else {
                        this.rooms.push({ id: Date.now(), ...this.formRoom });
                        this.showToast('Ruangan baru ditambahkan');
                    }

                    this.rooms = [...this.rooms];
                    localStorage.setItem('sarpras_sr_rooms', JSON.stringify(this.rooms));
                    this.showRoomModal = false;
                    this.isEditingRoom = false;
                },

                confirmDeleteRoom(room) {
                    if (typeof room === 'object' && room !== null) {
                        this.roomToDelete = room;
                        this.roomToDeleteId = room.id || room.name;
                    } else {
                        this.roomToDelete = null;
                        this.roomToDeleteId = room;
                    }
                    this.showDeleteRoomModal = true;
                },

                executeDeleteRoom() {
                    if (this.roomToDelete) {
                        const idx = this.rooms.indexOf(this.roomToDelete);
                        if (idx !== -1) {
                            this.rooms.splice(idx, 1);
                        } else {
                            const targetId = String(this.roomToDeleteId || '');
                            this.rooms = this.rooms.filter(r => String(r.id || '') !== targetId && r.name !== targetId);
                        }
                    } else if (this.roomToDeleteId !== null && this.roomToDeleteId !== undefined) {
                        const targetId = String(this.roomToDeleteId);
                        this.rooms = this.rooms.filter(r => String(r.id || '') !== targetId && r.name !== targetId);
                    }

                    this.rooms = [...this.rooms];
                    localStorage.setItem('sarpras_sr_rooms', JSON.stringify(this.rooms));
                    this.showToast('Ruangan berhasil dihapus', 'error');

                    this.showDeleteRoomModal = false;
                    this.roomToDeleteId = null;
                    this.roomToDelete = null;
                },

                deleteBorrowing(b) {
                    if (typeof b === 'object' && b !== null) {
                        const idx = this.borrowings.indexOf(b);
                        if (idx !== -1) {
                            this.borrowings.splice(idx, 1);
                        } else {
                            const targetId = String(b.id);
                            this.borrowings = this.borrowings.filter(item => String(item.id) !== targetId);
                        }
                    } else {
                        const targetId = String(b);
                        this.borrowings = this.borrowings.filter(item => String(item.id) !== targetId);
                    }
                    this.borrowings = [...this.borrowings];
                    localStorage.setItem('sarpras_sr_borrowings', JSON.stringify(this.borrowings));
                    this.showToast('Catatan peminjaman berhasil dihapus', 'error');
                },

                deleteMaintenance(m) {
                    if (typeof m === 'object' && m !== null) {
                        const idx = this.maintenances.indexOf(m);
                        if (idx !== -1) {
                            this.maintenances.splice(idx, 1);
                        } else {
                            const targetId = String(m.id);
                            this.maintenances = this.maintenances.filter(item => String(item.id) !== targetId);
                        }
                    } else {
                        const targetId = String(m);
                        this.maintenances = this.maintenances.filter(item => String(item.id) !== targetId);
                    }
                    this.maintenances = [...this.maintenances];
                    localStorage.setItem('sarpras_sr_maintenances', JSON.stringify(this.maintenances));
                    this.showToast('Laporan perbaikan berhasil dihapus', 'error');
                },

                filterByRoom(roomName) {
                    this.currentTab = 'inventory';
                    this.invSearch = roomName;
                },

                openSchoolModal() {
                    this.formSchool = { ...this.schoolProfile };
                    this.showSchoolModal = true;
                    this.showUserDropdown = false;
                },

                onLogoSelected(event) {
                    const file = event.target.files[0];
                    if (!file) return;
                    if (file.size > 2 * 1024 * 1024) {
                        this.showToast('Ukuran gambar terlalu besar (maksimal 2MB)', 'error');
                        return;
                    }
                    const reader = new FileReader();
                    reader.onload = (e) => {
                        this.formSchool.logoUrl = e.target.result;
                        this.showToast('Foto logo berhasil dimuat');
                    };
                    reader.readAsDataURL(file);
                },

                removeLogo() {
                    this.formSchool.logoUrl = null;
                    this.schoolProfile.logoUrl = null;
                    localStorage.setItem('sarpras_sr_school_profile', JSON.stringify(this.schoolProfile));
                    this.showToast('Logo dikembalikan ke default', 'error');
                },

                saveSchoolProfile() {
                    this.schoolProfile = { ...this.formSchool };
                    this.showSchoolModal = false;
                    this.showToast('Profil & Logo Sekolah diperbarui');
                },

                openAccountModal() {
                    this.formUser = { ...this.userProfile };
                    this.showAccountModal = true;
                    this.showUserDropdown = false;
                },

                openBackupModal() {
                    this.showBackupModal = true;
                    this.showUserDropdown = false;
                },

                exportFullJSON() {
                    const backupData = {
                        appName: 'SARPRAS_SEKOLAH_RAKYAT_KEMENSOS',
                        exportedAt: new Date().toISOString(),
                        schoolProfile: this.schoolProfile,
                        userProfile: this.userProfile,
                        inventory: this.inventory,
                        rooms: this.rooms,
                        borrowings: this.borrowings,
                        maintenances: this.maintenances,
                        categories: this.categories
                    };
                    const jsonStr = JSON.stringify(backupData, null, 2);
                    const blob = new Blob([jsonStr], { type: 'application/json' });
                    const link = document.createElement('a');
                    const dateTag = new Date().toISOString().split('T')[0];
                    link.href = URL.createObjectURL(blob);
                    link.setAttribute('download', `BACKUP_SARPRAS_${dateTag}.json`);
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    this.showToast('Backup data lengkap (JSON) berhasil di-download!');
                },

                importJSON(event) {
                    const file = event.target.files[0];
                    if (!file) return;

                    const reader = new FileReader();
                    reader.onload = (e) => {
                        try {
                            const data = JSON.parse(e.target.result);
                            if (!data || typeof data !== 'object') throw new Error('Format file tidak valid');

                            if (data.schoolProfile) this.schoolProfile = { ...defaultSchoolProfile, ...data.schoolProfile };
                            if (data.userProfile) this.userProfile = { ...defaultUserProfile, ...data.userProfile };
                            if (Array.isArray(data.inventory)) this.inventory = data.inventory;
                            if (Array.isArray(data.rooms)) this.rooms = data.rooms;
                            if (Array.isArray(data.borrowings)) this.borrowings = data.borrowings;
                            if (Array.isArray(data.maintenances)) this.maintenances = data.maintenances;
                            if (Array.isArray(data.categories)) this.categories = data.categories;

                            localStorage.setItem('sarpras_sr_school_profile', JSON.stringify(this.schoolProfile));
                            localStorage.setItem('sarpras_sr_user_profile', JSON.stringify(this.userProfile));
                            localStorage.setItem('sarpras_sr_inventory', JSON.stringify(this.inventory));
                            localStorage.setItem('sarpras_sr_rooms', JSON.stringify(this.rooms));
                            localStorage.setItem('sarpras_sr_borrowings', JSON.stringify(this.borrowings));
                            localStorage.setItem('sarpras_sr_maintenances', JSON.stringify(this.maintenances));
                            localStorage.setItem('sarpras_sr_categories', JSON.stringify(this.categories));

                            this.showToast('Data berhasil dipulihkan dari file backup!');
                            this.showBackupModal = false;
                            this.$nextTick(() => this.updateCharts());
                        } catch (err) {
                            this.showToast('Gagal memuat file backup: ' + err.message, 'error');
                        }
                    };
                    reader.readAsText(file);
                    event.target.value = '';
                },

                saveUserProfile() {
                    if (this.formUser.name) {
                        const parts = this.formUser.name.trim().split(' ');
                        this.formUser.avatarText = parts.length >= 2 ? (parts[0][0] + parts[1][0]).toUpperCase() : parts[0].substring(0, 2).toUpperCase();
                    }
                    this.userProfile = { ...this.formUser };
                    this.showAccountModal = false;
                    this.showToast('Profil Operator diperbarui');
                },

                openClearModal() {
                    this.showClearModal = true;
                    this.showUserDropdown = false;
                },

                executeClearAllData() {
                    this.inventory.splice(0, this.inventory.length);
                    this.borrowings.splice(0, this.borrowings.length);
                    this.maintenances.splice(0, this.maintenances.length);
                    this.rooms.splice(0, this.rooms.length);
                    localStorage.removeItem('sarpras_sr_inventory');
                    localStorage.removeItem('sarpras_sr_borrowings');
                    localStorage.removeItem('sarpras_sr_maintenances');
                    localStorage.removeItem('sarpras_sr_rooms');
                    this.showClearModal = false;
                    this.showToast('Seluruh data inventaris telah dikosongkan!');
                    this.$nextTick(() => this.updateCharts());
                },

                renderCharts() {
                    const ctxCond = document.getElementById('conditionChart');
                    const ctxCat = document.getElementById('categoryChart');
                    if (!ctxCond || !ctxCat) return;

                    if (this.conditionChartInstance) this.conditionChartInstance.destroy();
                    if (this.categoryChartInstance) this.categoryChartInstance.destroy();

                    const baik = this.inventory.filter(i => i.condition === 'Baik').reduce((sum, i) => sum + Number(i.qty || 0), 0);
                    const ringan = this.inventory.filter(i => i.condition === 'Rusak Ringan').reduce((sum, i) => sum + Number(i.qty || 0), 0);
                    const berat = this.inventory.filter(i => i.condition === 'Rusak Berat').reduce((sum, i) => sum + Number(i.qty || 0), 0);

                    if (this.inventory.length > 0) {
                        this.conditionChartInstance = new Chart(ctxCond, {
                            type: 'doughnut',
                            data: {
                                labels: ['Baik', 'Rusak Ringan', 'Rusak Berat'],
                                datasets: [{
                                    data: [baik, ringan, berat],
                                    backgroundColor: ['#10b981', '#f59e0b', '#f43f5e']
                                }]
                            },
                            options: { responsive: true, maintainAspectRatio: false }
                        });

                        const catCounts = this.categories.map(c => {
                            return this.inventory.filter(i => i.category === c).reduce((sum, i) => sum + Number(i.qty || 0), 0);
                        });

                        this.categoryChartInstance = new Chart(ctxCat, {
                            type: 'bar',
                            data: {
                                labels: this.categories.map(c => c.split(' ')[0]),
                                datasets: [{
                                    data: catCounts,
                                    backgroundColor: '#be123c',
                                    borderRadius: 6
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } }
                            }
                        });
                    }
                },

                updateCharts() {
                    if (this.currentTab === 'dashboard') this.renderCharts();
                },

                printReport() {
                    window.print();
                }
            }
        }).mount('#app');
    </script>
</body>
</html>
