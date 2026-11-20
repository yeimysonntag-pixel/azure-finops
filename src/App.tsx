import React, { useState, useMemo, ChangeEvent, FormEvent, useEffect } from 'react';
import Papa from 'papaparse';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
} from 'chart.js';
import { Bar } from 'react-chartjs-2';

import type { CsvRow, CustomerGroup } from './utils/excelGenerator';
import {
  exportSingleCustomerExcel,
  exportAllCustomersZip
} from './utils/excelGenerator';
import { exportCustomerPdf } from './utils/pdfGenerator';

import {
  Upload,
  Download,
  FileText,
  Search,
  BarChart3,
  Building2,
  Layers,
  TrendingUp,
  CheckCircle2,
  Globe,
  LogOut,
  UserCheck,
  Lock,
  AlertCircle,
  KeyRound,
  UserPlus,
  Trash2,
  RefreshCw,
  Scissors
} from 'lucide-react';

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend);

interface AppUser {
  id: number;
  email: string;
  role: 'admin' | 'user';
  country: string;
}

const STORAGE_USERS_KEY = 'tivit_finops_users_v17';
const STORAGE_CSV_KEY = 'tivit_finops_csv_v17';
const STORAGE_FILES_KEY = 'tivit_finops_files_v17';

const DEFAULT_USERS: AppUser[] = [
  { id: 1, email: 'chile.mgr@tivit.cl', role: 'user', country: 'Chile' },
  { id: 2, email: 'colombia.mgr@tivit.com', role: 'user', country: 'Colombia' },
  { id: 3, email: 'peru.mgr@tivit.com', role: 'user', country: 'Perú' },
  { id: 4, email: 'lorena.romero@tivit.com', role: 'user', country: 'Todos' },
  { id: 5, email: 'nora.rivera@tivit.com', role: 'user', country: 'Todos' },
  { id: 6, email: 'jorge.quiroz@tivit.com', role: 'user', country: 'Todos' },
  { id: 7, email: 'marcelo.fernandez@tivit.com', role: 'user', country: 'Todos' },
  { id: 8, email: 'rafael.sarmiento@tivit.com', role: 'user', country: 'Todos' },
  { id: 9, email: 'johana.arroyave@tivit.com', role: 'user', country: 'Todos' },
  { id: 10, email: 'danny.tambo@tivit.com', role: 'user', country: 'Todos' },
  { id: 11, email: 'fabio.castellanos@tivit.com', role: 'user', country: 'Todos' },
  { id: 12, email: 'darwin.jimenez@tivit.com', role: 'user', country: 'Todos' },
  { id: 13, email: 'diego.mejiaa@tivit.com', role: 'user', country: 'Todos' },
  { id: 14, email: 'roger.ungaro@tivit.com', role: 'user', country: 'Todos' }
];

const detectCountryByFileNameOrRow = (fileName: string, row: any): string => {
  const cleanFileName = fileName.toLowerCase();

  if (cleanFileName.includes('colombia')) return 'Colombia';
  if (cleanFileName.includes('peru')) return 'Perú';
  if (cleanFileName.includes('chile')) return 'Chile';
  if (cleanFileName.includes('mexico')) return 'México';

  const rowText = JSON.stringify(row).toLowerCase();
  if (/colombia|\bco\b|bogota/i.test(rowText)) return 'Colombia';
  if (/peru|perú|\bpe\b|lima/i.test(rowText)) return 'Perú';
  if (/chile|\bcl\b|santiago/i.test(rowText)) return 'Chile';
  if (/mexico|méxico|\bmx\b/i.test(rowText)) return 'México';

  return 'Colombia';
};

const parseScientificCost = (value: any): number => {
  if (value === undefined || value === null || value === '') return 0;
  const cleanStr = String(value).replace(',', '.').trim();
  const num = Number(cleanStr);
  return isNaN(num) ? 0 : num;
};

const extractRealCustomerName = (row: any): string => {
  if (row.customerName && String(row.customerName).trim() !== '') {
    return String(row.customerName).trim();
  }
  if (row.invoiceSectionName && String(row.invoiceSectionName).trim() !== '') {
    return String(row.invoiceSectionName).trim();
  }

  const keys = Object.keys(row);
  const candidates = ['subscriptionName', 'accountName', 'resourceGroup', 'partnerTenantTitle'];
  for (const candidate of candidates) {
    const matchedKey = keys.find(k => k.trim().toLowerCase() === candidate.toLowerCase());
    if (matchedKey && row[matchedKey] && String(row[matchedKey]).trim() !== '') {
      return String(row[matchedKey]).trim();
    }
  }

  return 'Cliente_General';
};

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [loginEmailInput, setLoginEmailInput] = useState<string>('');
  const [adminPasswordInput, setAdminPasswordInput] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string>('');

  const ADMIN_PASSWORD_SECRET = 'Tivit2026!';
  const PRIMARY_ADMIN_EMAIL = 'yeimy.sonntag@tivit.com';

  const [currentUser, setCurrentUser] = useState<AppUser>({
    id: 0,
    email: PRIMARY_ADMIN_EMAIL,
    role: 'admin',
    country: 'Todos'
  });

  const [usersList, setUsersList] = useState<AppUser[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_USERS_KEY);
      return saved ? JSON.parse(saved) : DEFAULT_USERS;
    } catch {
      return DEFAULT_USERS;
    }
  });

  const [csvRows, setCsvRows] = useState<CsvRow[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_CSV_KEY);
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [loadedFiles, setLoadedFiles] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_FILES_KEY);
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [selectedCostKey] = useState<string>('paygCostInUsd');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedCountryFilter, setSelectedCountryFilter] = useState<string>('Todos');
  const [selectedCustomer, setSelectedCustomer] = useState<CustomerGroup | null>(null);

  const [newUserEmail, setNewUserEmail] = useState<string>('');
  const [newUserRole, setNewUserRole] = useState<'admin' | 'user'>('user');
  const [newUserCountry, setNewUserCountry] = useState<string>('Todos');

  useEffect(() => {
    localStorage.setItem(STORAGE_USERS_KEY, JSON.stringify(usersList));
  }, [usersList]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_CSV_KEY, JSON.stringify(csvRows));
      localStorage.setItem(STORAGE_FILES_KEY, JSON.stringify(loadedFiles));
    } catch (e) {
      console.warn('Memoria local satisfecha.');
    }
  }, [csvRows, loadedFiles]);

  const handleEmergencyPurge = () => {
    if (window.confirm('¿Deseas purgar toda la memoria local y recargar la aplicación completamente limpia?')) {
      localStorage.clear();
      sessionStorage.clear();
      window.location.reload();
    }
  };

  useEffect(() => {
    const fetchRepoCsvFiles = async () => {
      if (!isAuthenticated || csvRows.length > 0) return;

      const filesToFetch = [
        './data/consumos_colombia.csv',
        './data/consumos_chile.csv',
        './data/consumos_peru.csv'
      ];

      let combinedRows: CsvRow[] = [];
      let loadedNames: string[] = [];

      for (const filePath of filesToFetch) {
        try {
          const res = await fetch(`${filePath}?t=${Date.now()}`);
          if (res.ok) {
            const csvText = await res.text();
            const fileName = filePath.split('/').pop() || 'dataset.csv';

            Papa.parse<CsvRow>(csvText, {
              header: true,
              skipEmptyLines: true,
              complete: (result) => {
                const rowsWithMeta = result.data.map((r: any) => {
                  const custName = extractRealCustomerName(r);
                  const detectedCountry = detectCountryByFileNameOrRow(fileName, r);

                  return {
                    ...r,
                    customerName: custName,
                    invoiceSectionName: r.invoiceSectionName || custName,
                    _sourceFileName: fileName,
                    _detectedCountry: detectedCountry
                  };
                });

                combinedRows = [...combinedRows, ...rowsWithMeta];
                loadedNames.push(fileName);
              }
            });
          }
        } catch (e) {
          // Archivo opcional aún no publicado
        }
      }

      if (combinedRows.length > 0) {
        setCsvRows(combinedRows);
        setLoadedFiles(loadedNames);
      }
    };

    fetchRepoCsvFiles();
  }, [isAuthenticated]);

  const handleLogin = (e: FormEvent) => {
    e.preventDefault();
    setErrorMessage('');
    const cleanEmail = loginEmailInput.trim().toLowerCase();

    if (!cleanEmail.endsWith('@tivit.com') && !cleanEmail.endsWith('@tivit.cl')) {
      setErrorMessage('Acceso denegado: Debe utilizar su correo corporativo @tivit.com o @tivit.cl');
      return;
    }

    if (cleanEmail === PRIMARY_ADMIN_EMAIL) {
      if (adminPasswordInput !== ADMIN_PASSWORD_SECRET) {
        setErrorMessage('Contraseña de administrador incorrecta.');
        return;
      }
      setCurrentUser({ id: 0, email: cleanEmail, role: 'admin', country: 'Todos' });
      setIsAuthenticated(true);
      return;
    }

    const registeredUser = usersList.find((u) => u.email.toLowerCase() === cleanEmail);
    if (!registeredUser) {
      setErrorMessage('Acceso denegado: Este correo no está autorizado por el Administrador.');
      return;
    }

    if (registeredUser.role === 'admin' && adminPasswordInput !== ADMIN_PASSWORD_SECRET) {
      setErrorMessage('Contraseña de administrador incorrecta.');
      return;
    }

    setCurrentUser({
      id: registeredUser.id,
      email: cleanEmail,
      role: registeredUser.role,
      country: registeredUser.role === 'admin' ? 'Todos' : registeredUser.country
    });
    setIsAuthenticated(true);
  };

  const handleLogout = () => {
    setIsAuthenticated(false);
    setLoginEmailInput('');
    setAdminPasswordInput('');
    setErrorMessage('');
  };

  const handleFileUploadAndMapping = (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];

    Papa.parse<CsvRow>(file, {
      header: true,
      skipEmptyLines: true,
      complete: (result) => {
        const rawData = result.data;
        if (!rawData || rawData.length === 0) return;

        const rowsByCountry: Record<string, any[]> = {
          Chile: [],
          Colombia: [],
          Perú: []
        };

        const processedRows: CsvRow[] = rawData.map((row: any) => {
          const rawCustomerName = extractRealCustomerName(row);
          const detectedCountry = detectCountryByFileNameOrRow(file.name, row);

          const cleanRow = {
            customerName: rawCustomerName,
            invoiceSectionName: row.invoiceSectionName || rawCustomerName,
            meterId: row.meterId || '',
            chargeType: row.chargeType || 'Usage',
            meterName: row.meterName || '',
            meterCategory: row.meterCategory || 'Otros',
            meterSubCategory: row.meterSubCategory || '',
            serviceFamily: row.serviceFamily || '',
            product: row.product || '',
            payGPrice: String(row.payGPrice || '0'),
            date: row.date || '',
            unitOfMeasure: row.unitOfMeasure || '',
            quantity: String(row.quantity || '0'),
            paygCostInUsd: String(row.paygCostInUsd || row.costInUsd || row.Cost || 0)
          };

          if (rowsByCountry[detectedCountry]) {
            rowsByCountry[detectedCountry].push(cleanRow);
          } else {
            rowsByCountry['Colombia'].push(cleanRow);
          }

          return {
            ...cleanRow,
            _sourceFileName: file.name,
            _detectedCountry: detectedCountry
          };
        });

        Object.entries(rowsByCountry).forEach(([country, rows]) => {
          if (rows.length > 0) {
            const csvOutput = Papa.unparse(rows);
            const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.setAttribute('href', url);
            link.setAttribute('download', `consumos_${country.toLowerCase().replace('ú', 'u')}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
          }
        });

        setCsvRows(processedRows);
        setLoadedFiles([file.name]);
        setSelectedCustomer(null);

        alert('¡CSV procesado con éxito! Se han descargado automáticamente los archivos livianos por país (consumos_colombia.csv, etc.) para que los coloques en public/data/.');
      }
    });
  };

  const periodDetected = useMemo(() => {
    if (!csvRows.length) return '202609';
    const sample = csvRows.find((r) => r.date)?.date;
    if (!sample) return '202609';
    const parts = String(sample).split('/');
    if (parts.length === 3) {
      return `${parts[2]}${parts[1]}`;
    }
    return String(sample).replace(/[-/]/g, '').substring(0, 6);
  }, [csvRows]);

  const { groupedByCountry, allCustomersList } = useMemo(() => {
    const countryMap: Record<string, Record<string, CustomerGroup>> = {};
    const list: CustomerGroup[] = [];

    csvRows.forEach((row: any) => {
      const custName = row.customerName || extractRealCustomerName(row);
      const country = row._detectedCountry || detectCountryByFileNameOrRow(row._sourceFileName || '', row);
      const cost = parseScientificCost(row.paygCostInUsd);

      if (!countryMap[country]) countryMap[country] = {};

      if (!countryMap[country][custName]) {
        const newGroup: CustomerGroup = {
          name: custName,
          country,
          totalCost: 0,
          records: [],
          categoryTotals: {}
        };
        countryMap[country][custName] = newGroup;
        list.push(newGroup);
      }

      const group = countryMap[country][custName];
      group.totalCost += cost;
      group.records.push(row);

      const category = row.meterCategory || 'Otros';
      group.categoryTotals[category] = (group.categoryTotals[category] || 0) + cost;
    });

    return { groupedByCountry: countryMap, allCustomersList: list };
  }, [csvRows]);

  const visibleCustomers = useMemo(() => {
    return allCustomersList.filter((cust) => {
      if (currentUser.role === 'user' && currentUser.country !== 'Todos') {
        if (cust.country !== currentUser.country) return false;
      }
      if ((currentUser.role === 'admin' || currentUser.country === 'Todos') && selectedCountryFilter !== 'Todos') {
        if (cust.country !== selectedCountryFilter) return false;
      }
      if (searchTerm && !cust.name.toLowerCase().includes(searchTerm.toLowerCase())) {
        return false;
      }
      return true;
    });
  }, [allCustomersList, selectedCountryFilter, searchTerm, currentUser]);

  const currentViewTotalCost = useMemo(() => {
    return visibleCustomers.reduce((acc, curr) => acc + curr.totalCost, 0);
  }, [visibleCustomers]);

  useEffect(() => {
    if (visibleCustomers.length > 0) {
      if (!selectedCustomer || !visibleCustomers.some((c) => c.name === selectedCustomer.name && c.country === selectedCustomer.country)) {
        setSelectedCustomer(visibleCustomers[0]);
      }
    } else {
      setSelectedCustomer(null);
    }
  }, [visibleCustomers]);

  const chartData = useMemo(() => {
    if (!selectedCustomer) return null;
    const sorted = Object.entries(selectedCustomer.categoryTotals).sort((a, b) => b[1] - a[1]);
    return {
      labels: sorted.map((e) => e[0]),
      datasets: [
        {
          label: `Consumo USD (${selectedCostKey})`,
          data: sorted.map((e) => e[1]),
          backgroundColor: '#0284c7',
          borderRadius: 6
        }
      ]
    };
  }, [selectedCustomer, selectedCostKey]);

  const handleAddUser = (e: FormEvent) => {
    e.preventDefault();
    const cleanNewEmail = newUserEmail.trim().toLowerCase();

    if (!cleanNewEmail.endsWith('@tivit.com') && !cleanNewEmail.endsWith('@tivit.cl')) {
      alert('Error: Solo se autorizan correos corporativos @tivit.com o @tivit.cl');
      return;
    }

    if (usersList.some((u) => u.email.toLowerCase() === cleanNewEmail)) {
      alert('El usuario ya está registrado.');
      return;
    }

    const newUser: AppUser = {
      id: Date.now(),
      email: cleanNewEmail,
      role: newUserRole,
      country: newUserRole === 'admin' ? 'Todos' : newUserCountry
    };

    setUsersList((prev) => [...prev, newUser]);
    setNewUserEmail('');
  };

  const handleToggleUserRole = (id: number) => {
    setUsersList((prev) =>
      prev.map((u) => {
        if (u.id === id) {
          const nextRole = u.role === 'admin' ? 'user' : 'admin';
          return { ...u, role: nextRole, country: nextRole === 'admin' ? 'Todos' : u.country };
        }
        return u;
      })
    );
  };

  const handleDeleteUser = (id: number, email: string) => {
    if (window.confirm(`¿Deseas revocar el acceso a ${email}?`)) {
      setUsersList((prev) => prev.filter((u) => u.id !== id));
    }
  };

  if (!isAuthenticated) {
    const cleanInputEmail = loginEmailInput.trim().toLowerCase();
    const isPrimaryAdmin = cleanInputEmail === PRIMARY_ADMIN_EMAIL;
    const registeredUserObj = usersList.find((u) => u.email.toLowerCase() === cleanInputEmail);
    const requiresPassword = isPrimaryAdmin || registeredUserObj?.role === 'admin';

    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4">
        <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-md space-y-6">
          <div className="text-center space-y-2">
            <div className="bg-sky-600 w-12 h-12 rounded-2xl flex items-center justify-center text-white mx-auto shadow-lg shadow-sky-600/30">
              <Lock className="w-6 h-6" />
            </div>
            <h1 className="text-xl font-extrabold text-slate-900">TIVIT Azure FinOps Portal</h1>
            <p className="text-xs text-slate-500">Acceso restringido a personal autorizado (@tivit.com / @tivit.cl)</p>
          </div>

          {errorMessage && (
            <div className="bg-red-50 border border-red-200 text-red-700 text-xs p-3 rounded-xl flex items-start space-x-2">
              <AlertCircle className="w-4 h-4 text-red-600 shrink-0 mt-0.5" />
              <span>{errorMessage}</span>
            </div>
          )}

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1">Correo Corporativo TIVIT</label>
              <input
                type="email"
                placeholder="usuario@tivit.com o usuario@tivit.cl"
                value={loginEmailInput}
                onChange={(e) => {
                  setLoginEmailInput(e.target.value);
                  setErrorMessage('');
                }}
                className="w-full border border-slate-300 p-3 rounded-xl text-xs focus:ring-2 focus:ring-sky-500 focus:outline-none"
                required
              />
            </div>

            {requiresPassword && (
              <div className="space-y-1">
                <label className="block text-xs font-bold text-slate-700 flex items-center space-x-1">
                  <KeyRound className="w-3.5 h-3.5 text-sky-600" />
                  <span>Contraseña Administrador</span>
                </label>
                <input
                  type="password"
                  placeholder="Ingrese su clave secreta"
                  value={adminPasswordInput}
                  onChange={(e) => setAdminPasswordInput(e.target.value)}
                  className="w-full border border-slate-300 p-3 rounded-xl text-xs focus:ring-2 focus:ring-sky-500 focus:outline-none"
                  required
                />
              </div>
            )}

            <button
              type="submit"
              className="w-full bg-sky-600 hover:bg-sky-700 text-white font-bold text-xs py-3 rounded-xl shadow-lg transition flex items-center justify-center space-x-2"
            >
              <UserCheck className="w-4 h-4" />
              <span>Ingresar a la Plataforma</span>
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800">
      <header className="bg-slate-900 text-white sticky top-0 z-50 border-b border-slate-800 shadow-lg">
        <div className="max-w-7xl mx-auto px-6 h-16 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <div className="bg-sky-600 p-2 rounded-xl text-white shadow-md">
              <TrendingUp className="w-5 h-5" />
            </div>
            <div>
              <h1 className="font-extrabold text-base tracking-tight">Azure FinOps Manager</h1>
              <p className="text-[10px] text-sky-400 font-semibold uppercase tracking-wider">
                {currentUser.role === 'admin' || currentUser.country === 'Todos'
                  ? 'Gestión Consolidada Multipaís'
                  : `Vista Asignada: ${currentUser.country}`}
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-4 text-xs">
            <div className="bg-slate-800 px-3.5 py-1.5 rounded-full border border-slate-700 flex items-center space-x-2">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
              <span className="text-slate-300 font-medium">{currentUser.email}</span>
              <span className="bg-sky-500/20 text-sky-300 text-[10px] font-bold px-2 py-0.5 rounded-full uppercase">
                {currentUser.role === 'admin' ? 'Administrador' : `Auditor ${currentUser.country}`}
              </span>
            </div>

            <button
              onClick={handleEmergencyPurge}
              className="bg-red-600 hover:bg-red-700 text-white px-3 py-1.5 rounded-lg font-bold transition flex items-center space-x-1"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Purga Total</span>
            </button>

            <button
              onClick={handleLogout}
              className="text-slate-400 hover:text-white border border-slate-700 hover:bg-slate-800 px-3 py-1.5 rounded-lg font-medium transition flex items-center space-x-1.5"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span>Cerrar Sesión</span>
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-8 space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          {currentUser.role === 'admin' ? (
            <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm col-span-1 md:col-span-2 flex flex-col justify-between">
              <div>
                <h2 className="text-sm font-bold text-slate-900 flex items-center space-x-2">
                  <Scissors className="w-4 h-4 text-sky-600" />
                  <span>Procesar y Fragmentar CSV Masivo</span>
                </h2>
                <p className="text-xs text-slate-500 mt-1">Sube el CSV masivo de Azure. Se limpiará y descargará automáticamente separado por país.</p>
              </div>

              <div className="mt-4 flex flex-wrap items-center gap-3">
                <label className="bg-sky-600 hover:bg-sky-700 text-white text-xs font-bold px-4 py-2.5 rounded-xl cursor-pointer transition shadow flex items-center space-x-2">
                  <Upload className="w-4 h-4" />
                  <span>Seleccionar CSV Pesado</span>
                  <input type="file" accept=".csv" onChange={handleFileUploadAndMapping} className="hidden" />
                </label>

                {loadedFiles.length > 0 && (
                  <div className="flex items-center space-x-2">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                    <span className="text-xs font-bold text-slate-700">
                      {csvRows.length.toLocaleString()} filas procesadas
                    </span>
                  </div>
                )}
              </div>
            </div>
          ) : (
            <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm col-span-1 md:col-span-2 flex items-center space-x-4">
              <div className="bg-sky-50 p-3 rounded-2xl text-sky-600">
                <Globe className="w-6 h-6" />
              </div>
              <div>
                <h2 className="text-sm font-bold text-slate-900">Operación Asignada: {currentUser.country}</h2>
                <p className="text-xs text-slate-500 mt-0.5">Acceso a informes e indicadores de {currentUser.country}.</p>
              </div>
            </div>
          )}

          <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm">
            <p className="text-xs font-semibold text-slate-500">
              {currentUser.role === 'admin' || currentUser.country === 'Todos'
                ? 'Costo Total Consolidado'
                : `Costo Total (${currentUser.country})`}
            </p>
            <p className="text-2xl font-black text-slate-900 mt-1">${currentViewTotalCost.toFixed(6)} USD</p>
            <p className="text-[10px] font-bold text-sky-600 mt-1">Periodo Procesado: {periodDetected}</p>
          </div>

          <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm flex flex-col justify-between">
            <div>
              <p className="text-xs font-semibold text-slate-500">
                {currentUser.role === 'admin' || currentUser.country === 'Todos'
                  ? 'Clientes Visibles'
                  : `Clientes en ${currentUser.country}`}
              </p>
              <p className="text-2xl font-black text-slate-900 mt-1">{visibleCustomers.length}</p>
            </div>
            {visibleCustomers.length > 0 && (
              <button
                onClick={() => exportAllCustomersZip(visibleCustomers, periodDetected, selectedCostKey)}
                className="w-full bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold py-2 rounded-xl transition shadow flex items-center justify-center space-x-1 mt-2"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Descargar ZIP ({currentUser.country})</span>
              </button>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-4 bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm space-y-4">
            <div className="flex justify-between items-center">
              <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center space-x-2">
                <Building2 className="w-4 h-4 text-sky-600" />
                <span>Clientes ({visibleCustomers.length})</span>
              </h3>
              <div className="relative w-36">
                <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="Buscar..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-8 pr-2 py-1 bg-slate-50 border rounded-lg text-xs"
                />
              </div>
            </div>

            {currentUser.role === 'admin' || currentUser.country === 'Todos' ? (
              <div className="flex flex-wrap gap-1">
                {['Todos', ...Object.keys(groupedByCountry)].map((country) => (
                  <button
                    key={country}
                    onClick={() => setSelectedCountryFilter(country)}
                    className={`text-[11px] px-2.5 py-1 rounded-lg font-semibold transition ${
                      selectedCountryFilter === country
                        ? 'bg-sky-600 text-white shadow-sm'
                        : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                    }`}
                  >
                    {country}
                  </button>
                ))}
              </div>
            ) : (
              <div className="flex items-center space-x-1.5 text-xs text-sky-700 bg-sky-50 px-3 py-1.5 rounded-lg border border-sky-100 font-bold">
                <Globe className="w-3.5 h-3.5" />
                <span>Filtro Activo: {currentUser.country}</span>
              </div>
            )}

            <div className="space-y-2 max-h-[380px] overflow-y-auto pr-1 custom-scrollbar">
              {visibleCustomers.length === 0 ? (
                <p className="text-xs text-slate-400 text-center py-8">
                  {csvRows.length === 0 ? 'Cargue un CSV para visualizar métricas' : `Sin datos registrados para ${currentUser.country}`}
                </p>
              ) : (
                visibleCustomers.map((cust) => (
                  <div
                    key={`${cust.country}_${cust.name}`}
                    onClick={() => setSelectedCustomer(cust)}
                    className={`p-3 rounded-xl border transition cursor-pointer flex justify-between items-center ${
                      selectedCustomer?.name === cust.name && selectedCustomer?.country === cust.country
                        ? 'bg-sky-50 border-sky-500 ring-1 ring-sky-500 shadow-sm'
                        : 'bg-white border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    <div>
                      <p className="text-xs font-bold text-slate-900 truncate max-w-[150px]">{cust.name}</p>
                      <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded">
                        {cust.country}
                      </span>
                    </div>
                    <div className="text-right">
                      <p className="text-xs font-extrabold text-sky-700">${cust.totalCost.toFixed(4)}</p>
                      <p className="text-[9px] text-slate-400">USD</p>
                    </div>
                  </div>
                ))
              )}
            </div>

            {currentUser.role === 'admin' && (
              <div className="border-t pt-4 space-y-3">
                <h4 className="text-xs font-bold text-slate-800 flex items-center space-x-1.5">
                  <UserPlus className="w-4 h-4 text-emerald-600" />
                  <span>Asignar Permisos & Roles</span>
                </h4>

                <form onSubmit={handleAddUser} className="space-y-2">
                  <input
                    type="email"
                    placeholder="correo@tivit.com o @tivit.cl"
                    value={newUserEmail}
                    onChange={(e) => setNewUserEmail(e.target.value)}
                    className="w-full border p-2 text-xs rounded-lg"
                    required
                  />

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="text-[10px] font-bold text-slate-500 block mb-0.5">Rol de Usuario</label>
                      <select
                        value={newUserRole}
                        onChange={(e) => setNewUserRole(e.target.value as 'admin' | 'user')}
                        className="w-full border p-2 text-xs rounded-lg bg-slate-50"
                      >
                        <option value="user">Usuario / Auditor</option>
                        <option value="admin">Administrador</option>
                      </select>
                    </div>

                    <div>
                      <label className="text-[10px] font-bold text-slate-500 block mb-0.5">País Operación</label>
                      <select
                        value={newUserCountry}
                        onChange={(e) => setNewUserCountry(e.target.value)}
                        disabled={newUserRole === 'admin'}
                        className="w-full border p-2 text-xs rounded-lg bg-slate-50 disabled:opacity-50"
                      >
                        <option value="Todos">Todos los países</option>
                        <option value="Chile">Chile</option>
                        <option value="Colombia">Colombia</option>
                        <option value="Perú">Perú</option>
                        <option value="México">México</option>
                      </select>
                    </div>
                  </div>

                  <button type="submit" className="w-full bg-emerald-600 hover:bg-emerald-700 text-white text-xs py-2 rounded-lg font-bold transition">
                    Guardar y Asignar Rol
                  </button>
                </form>

                <div className="space-y-1.5 pt-2">
                  <p className="text-[10px] font-bold text-slate-400 uppercase">Usuarios Autorizados ({usersList.length})</p>
                  <div className="max-h-28 overflow-y-auto space-y-1 custom-scrollbar">
                    {usersList.map((user) => (
                      <div key={user.id} className="flex items-center justify-between bg-slate-50 p-2 rounded-lg text-[11px]">
                        <div className="truncate max-w-[150px]">
                          <p className="font-bold text-slate-800 truncate">{user.email}</p>
                          <span className="text-[9px] text-slate-500">{user.role === 'admin' ? 'Acceso Global Admin' : `Auditor: ${user.country}`}</span>
                        </div>

                        <div className="flex items-center space-x-1">
                          <button
                            onClick={() => handleToggleUserRole(user.id)}
                            className="text-[10px] font-bold bg-slate-200 hover:bg-sky-100 text-slate-700 hover:text-sky-800 px-2 py-1 rounded transition"
                          >
                            {user.role === 'admin' ? 'Hacer Usuario' : 'Hacer Admin'}
                          </button>

                          <button
                            onClick={() => handleDeleteUser(user.id, user.email)}
                            className="text-red-500 hover:text-red-700 hover:bg-red-50 p-1 rounded transition"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>

          <div className="lg:col-span-8 bg-white p-6 rounded-2xl border border-slate-200/80 shadow-sm space-y-6">
            {selectedCustomer ? (
              <>
                <div className="flex flex-col sm:flex-row sm:items-center justify-between border-b pb-4 gap-4">
                  <div>
                    <span className="text-[10px] font-bold bg-sky-100 text-sky-800 px-2.5 py-1 rounded-full uppercase">
                      {selectedCustomer.country}
                    </span>
                    <h2 className="text-xl font-extrabold text-slate-900 mt-1">{selectedCustomer.name}</h2>
                  </div>

                  <div className="flex items-center space-x-3">
                    <div className="text-right mr-2">
                      <p className="text-xs text-slate-500 font-medium">Consumo Total</p>
                      <p className="text-2xl font-black text-sky-700">${selectedCustomer.totalCost.toFixed(6)} USD</p>
                    </div>

                    <button
                      onClick={() => exportSingleCustomerExcel(selectedCustomer, periodDetected, selectedCostKey)}
                      className="bg-sky-600 hover:bg-sky-700 text-white text-xs font-bold px-3.5 py-2.5 rounded-xl shadow transition flex items-center space-x-1.5"
                    >
                      <Download className="w-4 h-4" />
                      <span>Excel</span>
                    </button>

                    <button
                      onClick={() => exportCustomerPdf(selectedCustomer, periodDetected)}
                      className="bg-red-600 hover:bg-red-700 text-white text-xs font-bold px-3.5 py-2.5 rounded-xl shadow transition flex items-center space-x-1.5"
                    >
                      <FileText className="w-4 h-4" />
                      <span>PDF</span>
                    </button>
                  </div>
                </div>

                {chartData && (
                  <div className="border border-slate-200/80 p-4 rounded-xl bg-slate-50/50 space-y-2">
                    <h4 className="text-xs font-bold text-slate-700 flex items-center space-x-1.5">
                      <BarChart3 className="w-4 h-4 text-sky-600" />
                      <span>Distribución por Categoría de Servicio</span>
                    </h4>
                    <div className="h-52">
                      <Bar data={chartData} options={{ responsive: true, maintainAspectRatio: false }} />
                    </div>
                  </div>
                )}

                <div className="space-y-2">
                  <h4 className="text-xs font-bold text-slate-700 flex items-center space-x-1.5">
                    <Layers className="w-4 h-4 text-sky-600" />
                    <span>Detalle de Consumos ({selectedCustomer.records.length} líneas)</span>
                  </h4>

                  <div className="overflow-x-auto border border-slate-200 rounded-xl max-h-72 custom-scrollbar">
                    <table className="w-full text-left text-[11px] text-slate-600">
                      <thead className="bg-slate-900 text-white font-semibold sticky top-0">
                        <tr>
                          <th className="p-2.5">invoiceSectionName</th>
                          <th className="p-2.5">meterId</th>
                          <th className="p-2.5">chargeType</th>
                          <th className="p-2.5">meterName</th>
                          <th className="p-2.5">meterCategory</th>
                          <th className="p-2.5">meterSubCategory</th>
                          <th className="p-2.5">serviceFamily</th>
                          <th className="p-2.5">product</th>
                          <th className="p-2.5">payGPrice</th>
                          <th className="p-2.5">date</th>
                          <th className="p-2.5">unitOfMeasure</th>
                          <th className="p-2.5">quantity</th>
                          <th className="p-2.5">Costo USD</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {selectedCustomer.records.map((r: any, i: number) => (
                          <tr key={i} className="hover:bg-slate-50 transition">
                            <td className="p-2.5 font-medium text-slate-900">{r.invoiceSectionName || selectedCustomer.name}</td>
                            <td className="p-2.5 font-mono text-[10px] text-slate-400 truncate max-w-[90px]" title={r.meterId}>{r.meterId}</td>
                            <td className="p-2.5">{r.chargeType}</td>
                            <td className="p-2.5 truncate max-w-[100px]" title={r.meterName}>{r.meterName}</td>
                            <td className="p-2.5 font-bold text-sky-700">{r.meterCategory}</td>
                            <td className="p-2.5">{r.meterSubCategory}</td>
                            <td className="p-2.5">{r.serviceFamily}</td>
                            <td className="p-2.5 truncate max-w-[100px]" title={r.product}>{r.product}</td>
                            <td className="p-2.5">${parseFloat(String(r.payGPrice || 0)).toFixed(4)}</td>
                            <td className="p-2.5 whitespace-nowrap">{r.date}</td>
                            <td className="p-2.5">{r.unitOfMeasure}</td>
                            <td className="p-2.5">{r.quantity}</td>
                            <td className="p-2.5 font-extrabold text-emerald-600 whitespace-nowrap">
                              ${parseScientificCost(r.paygCostInUsd).toFixed(6)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            ) : (
              <div className="text-center py-20 text-slate-400">
                <Building2 className="w-12 h-12 mx-auto mb-2 opacity-40" />
                <p className="text-sm font-medium">Selecciona un cliente para explorar sus métricas FinOps.</p>
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}