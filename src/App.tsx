import React, { useState, useMemo, ChangeEvent, FormEvent } from 'react';
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

// Importación estricta de tipos e interfaces
import type { CsvRow, CustomerGroup } from './utils/excelGenerator';
import {
  detectCountry,
  exportSingleCustomerExcel,
  exportAllCustomersZip
} from './utils/excelGenerator';
import { exportCustomerPdf } from './utils/pdfGenerator';

import {
  Upload,
  Download,
  FileSpreadsheet,
  FileText,
  Search,
  BarChart3,
  Building2,
  Layers,
  ShieldCheck,
  TrendingUp,
  CheckCircle2,
  Globe,
  LogOut,
  UserCheck,
  Lock,
  AlertCircle,
  KeyRound,
  UserPlus,
  Trash2
} from 'lucide-react';

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend);

interface AppUser {
  id: number;
  email: string;
  role: 'admin' | 'user';
  country: string;
}

export default function App() {
  // Estado para la autenticación y control de acceso
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [loginEmailInput, setLoginEmailInput] = useState<string>('');
  const [adminPasswordInput, setAdminPasswordInput] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string>('');

  // Clave del administrador principal
  const ADMIN_PASSWORD_SECRET = 'Tivit2026!';
  const PRIMARY_ADMIN_EMAIL = 'yeimy.sonntag@tivit.com';

  const [currentUser, setCurrentUser] = useState<AppUser>({
    id: 0,
    email: PRIMARY_ADMIN_EMAIL,
    role: 'admin',
    country: 'Todos'
  });

  // Lista dinámica de usuarios gestionada por el administrador
  const [usersList, setUsersList] = useState<AppUser[]>([
    { id: 1, email: 'chile.mgr@tivit.cl', role: 'user', country: 'Chile' },
    { id: 2, email: 'colombia.mgr@tivit.com', role: 'user', country: 'Colombia' },
    { id: 3, email: 'peru.mgr@tivit.com', role: 'user', country: 'Perú' }
  ]);

  const [csvRows, setCsvRows] = useState<CsvRow[]>([]);
  const [loadedFiles, setLoadedFiles] = useState<string[]>([]);
  const [selectedCostKey, setSelectedCostKey] = useState<string>('paygCostInUsd');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedCountryFilter, setSelectedCountryFilter] = useState<string>('Todos');
  const [selectedCustomer, setSelectedCustomer] = useState<CustomerGroup | null>(null);

  // Estados para el formulario de registro de nuevos usuarios/roles
  const [newUserEmail, setNewUserEmail] = useState<string>('');
  const [newUserRole, setNewUserRole] = useState<'admin' | 'user'>('user');
  const [newUserCountry, setNewUserCountry] = useState<string>('Chile');

  // AUTENTICACIÓN Y VALIDACIÓN DE ROLES
  const handleLogin = (e: FormEvent) => {
    e.preventDefault();
    setErrorMessage('');

    const cleanEmail = loginEmailInput.trim().toLowerCase();

    // 1. Validar dominio corporativo @tivit.com o @tivit.cl
    const hasValidDomain = cleanEmail.endsWith('@tivit.com') || cleanEmail.endsWith('@tivit.cl');
    if (!hasValidDomain) {
      setErrorMessage('Acceso denegado: El correo debe pertenecer al dominio @tivit.com o @tivit.cl');
      return;
    }

    // 2. Administrador Principal por Defecto
    if (cleanEmail === PRIMARY_ADMIN_EMAIL) {
      if (adminPasswordInput !== ADMIN_PASSWORD_SECRET) {
        setErrorMessage('Contraseña de administrador incorrecta.');
        return;
      }
      setCurrentUser({
        id: 0,
        email: cleanEmail,
        role: 'admin',
        country: 'Todos'
      });
      setIsAuthenticated(true);
      return;
    }

    // 3. Buscar en la lista de usuarios autorizados previamente por el admin
    const registeredUser = usersList.find(u => u.email.toLowerCase() === cleanEmail);
    if (!registeredUser) {
      setErrorMessage('Acceso denegado: Este correo no está autorizado por el Administrador.');
      return;
    }

    // Si el usuario registrado fue asignado como Administrador, pedir contraseña
    if (registeredUser.role === 'admin') {
      if (adminPasswordInput !== ADMIN_PASSWORD_SECRET) {
        setErrorMessage('Contraseña de administrador incorrecta.');
        return;
      }
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

  // CARGA MÚLTIPLE DE HASTA 3 ARCHIVOS CSV
  const handleMultipleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const fileList = Array.from(files).slice(0, 3);
    const fileNames = fileList.map(f => f.name);
    setLoadedFiles(fileNames);

    let combinedRows: CsvRow[] = [];
    let processedCount = 0;

    fileList.forEach(file => {
      Papa.parse<CsvRow>(file, {
        header: true,
        skipEmptyLines: true,
        complete: (result) => {
          const rowsWithFileRef = result.data.map(r => ({
            ...r,
            _sourceFileName: file.name
          }));

          combinedRows = [...combinedRows, ...rowsWithFileRef];
          processedCount++;

          if (processedCount === fileList.length) {
            setCsvRows(combinedRows);

            if (combinedRows.length > 0) {
              const firstRow = combinedRows[0];
              const keys = Object.keys(firstRow).filter(k =>
                k.toLowerCase().includes('cost') || k.toLowerCase().includes('price')
              );
              if (keys.includes('paygCostInUsd')) {
                setSelectedCostKey('paygCostInUsd');
              } else if (keys.length > 0) {
                setSelectedCostKey(keys[0]);
              }
            }
            setSelectedCustomer(null);
          }
        }
      });
    });
  };

  const periodDetected = useMemo(() => {
    if (!csvRows.length) return '202609';
    const sample = csvRows.find(r => r.billingPeriodEndDate || r.date)?.billingPeriodEndDate;
    if (sample) {
      const clean = String(sample).replace(/[-/]/g, '');
      return clean.substring(0, 6);
    }
    return '202609';
  }, [csvRows]);

  const { groupedByCountry, allCustomersList } = useMemo(() => {
    const countryMap: Record<string, Record<string, CustomerGroup>> = {};
    const list: CustomerGroup[] = [];

    csvRows.forEach(row => {
      const custName = row.customerName || row.invoiceSectionName || 'Cliente_General';
      const fileRef = (row as any)._sourceFileName || '';
      const country = detectCountry(`${custName}_${row.invoiceSectionName || ''}_${fileRef}`);
      const cost = parseFloat(String(row[selectedCostKey] || row.paygCostInUsd || row.costInUsd || 0));

      if (!countryMap[country]) {
        countryMap[country] = {};
      }

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
  }, [csvRows, selectedCostKey]);

  // FILTRADO SEGÚN ROL Y PAÍS PERMITIDO
  const visibleCustomers = useMemo(() => {
    return allCustomersList.filter(cust => {
      if (currentUser.role === 'user' && currentUser.country !== 'Todos') {
        if (cust.country !== currentUser.country) return false;
      }
      if (currentUser.role === 'admin' && selectedCountryFilter !== 'Todos') {
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

  React.useEffect(() => {
    if (visibleCustomers.length > 0) {
      if (!selectedCustomer || !visibleCustomers.some(c => c.name === selectedCustomer.name && c.country === selectedCustomer.country)) {
        setSelectedCustomer(visibleCustomers[0]);
      }
    } else {
      setSelectedCustomer(null);
    }
  }, [visibleCustomers]);

  const chartData = useMemo(() => {
    if (!selectedCustomer) return null;
    const sorted = Object.entries(selectedCustomer.categoryTotals).sort((a, b) => b[1] - a[1]);
    const labels = sorted.map(e => e[0]);
    const dataValues = sorted.map(e => e[1]);

    return {
      labels,
      datasets: [
        {
          label: `Consumo USD (${selectedCostKey})`,
          data: dataValues,
          backgroundColor: '#0284c7',
          borderRadius: 6
        }
      ]
    };
  }, [selectedCustomer, selectedCostKey]);

  // ASIGNACIÓN Y GESTIÓN DE ROLES/USUARIOS POR EL ADMINISTRADOR
  const handleAddUser = (e: FormEvent) => {
    e.preventDefault();
    const cleanNewEmail = newUserEmail.trim().toLowerCase();

    const hasValidDomain = cleanNewEmail.endsWith('@tivit.com') || cleanNewEmail.endsWith('@tivit.cl');
    if (!hasValidDomain) {
      alert('Error: Solo se pueden autorizar correos con dominio @tivit.com o @tivit.cl');
      return;
    }

    if (usersList.some(u => u.email.toLowerCase() === cleanNewEmail)) {
      alert('El usuario ya se encuentra en la lista. Puedes cambiar su rol o eliminarlo más abajo.');
      return;
    }

    const newUser: AppUser = {
      id: Date.now(),
      email: cleanNewEmail,
      role: newUserRole,
      country: newUserRole === 'admin' ? 'Todos' : newUserCountry
    };

    setUsersList([...usersList, newUser]);
    setNewUserEmail('');
  };

  // CAMBIAR ROL DE USUARIO
  const handleToggleUserRole = (id: number) => {
    setUsersList(usersList.map(u => {
      if (u.id === id) {
        const nextRole = u.role === 'admin' ? 'user' : 'admin';
        return {
          ...u,
          role: nextRole,
          country: nextRole === 'admin' ? 'Todos' : u.country
        };
      }
      return u;
    }));
  };

  // ELIMINAR USUARIO/REVOCAR ACCESO
  const handleDeleteUser = (id: number, email: string) => {
    if (window.confirm(`¿Estás seguro de revocar el acceso a ${email}?`)) {
      setUsersList(usersList.filter(u => u.id !== id));
    }
  };

  // MODAL LOGIN SEGURO CON DETECCIÓN DE ROL Y CLAVE
  if (!isAuthenticated) {
    const cleanInputEmail = loginEmailInput.trim().toLowerCase();
    const isPrimaryAdmin = cleanInputEmail === PRIMARY_ADMIN_EMAIL;
    const registeredUserObj = usersList.find(u => u.email.toLowerCase() === cleanInputEmail);
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
              <div className="space-y-1 animate-fadeIn">
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
              className="w-full bg-sky-600 hover:bg-sky-700 text-white font-bold text-xs py-3 rounded-xl shadow-lg shadow-sky-600/20 transition flex items-center justify-center space-x-2"
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
      {/* ENCABEZADO PRINCIPAL */}
      <header className="bg-slate-900 text-white sticky top-0 z-50 border-b border-slate-800 shadow-lg">
        <div className="max-w-7xl mx-auto px-6 h-16 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <div className="bg-sky-600 p-2 rounded-xl text-white shadow-md">
              <TrendingUp className="w-5 h-5" />
            </div>
            <div>
              <h1 className="font-extrabold text-base tracking-tight">Azure FinOps Manager</h1>
              <p className="text-[10px] text-sky-400 font-semibold uppercase tracking-wider">
                {currentUser.role === 'admin' ? 'Gestión Consolidada Multipaís' : `Vista Asignada: ${currentUser.country}`}
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-4 text-xs">
            <div className="bg-slate-800 px-3.5 py-1.5 rounded-full border border-slate-700/80 flex items-center space-x-2">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
              <span className="text-slate-300 font-medium">{currentUser.email}</span>
              <span className="bg-sky-500/20 text-sky-300 text-[10px] font-bold px-2 py-0.5 rounded-full uppercase">
                {currentUser.role === 'admin' ? 'Administrador' : `Auditor ${currentUser.country}`}
              </span>
            </div>

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

      {/* CONTENIDO PRINCIPAL */}
      <main className="max-w-7xl mx-auto px-6 py-8 space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          {currentUser.role === 'admin' ? (
            <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm col-span-1 md:col-span-2 flex flex-col justify-between">
              <div>
                <h2 className="text-sm font-bold text-slate-900 flex items-center space-x-2">
                  <Upload className="w-4 h-4 text-sky-600" />
                  <span>Cargar Consumos Azure (Hasta 3 CSVs)</span>
                </h2>
                <p className="text-xs text-slate-500 mt-1">Sube los extractos billing de la operación regional para actualizar datos.</p>
              </div>

              <div className="mt-4 flex flex-wrap items-center gap-3">
                <label className="bg-sky-600 hover:bg-sky-700 text-white text-xs font-bold px-4 py-2.5 rounded-xl cursor-pointer transition shadow-md shadow-sky-600/20 flex items-center space-x-2">
                  <FileSpreadsheet className="w-4 h-4" />
                  <span>Cargar CSVs</span>
                  <input type="file" accept=".csv" multiple onChange={handleMultipleFileUpload} className="hidden" />
                </label>

                {loadedFiles.length > 0 && (
                  <div className="flex items-center space-x-2">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                    <span className="text-xs font-bold text-slate-700">
                      {loadedFiles.length} archivo(s) ({csvRows.length.toLocaleString()} filas)
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
                <p className="text-xs text-slate-500 mt-0.5">Acceso restringido para auditoría y descarga de informes de {currentUser.country}.</p>
              </div>
            </div>
          )}

          <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm">
            <p className="text-xs font-semibold text-slate-500">
              {currentUser.role === 'admin' ? 'Costo Total Consolidado' : `Costo Total (${currentUser.country})`}
            </p>
            <p className="text-2xl font-black text-slate-900 mt-1">${currentViewTotalCost.toFixed(2)} USD</p>
            <p className="text-[10px] font-bold text-sky-600 mt-1">Periodo Procesado</p>
          </div>

          <div className="bg-white p-5 rounded-2xl border border-slate-200/80 shadow-sm flex flex-col justify-between">
            <div>
              <p className="text-xs font-semibold text-slate-500">
                {currentUser.role === 'admin' ? 'Clientes Visibles' : `Clientes en ${currentUser.country}`}
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

        {/* NAVEGACIÓN Y CLIENTES */}
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

            {currentUser.role === 'admin' ? (
              <div className="flex flex-wrap gap-1">
                {['Todos', ...Object.keys(groupedByCountry)].map(country => (
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
                <span>Filtro de Seguridad Activo: {currentUser.country}</span>
              </div>
            )}

            <div className="space-y-2 max-h-[380px] overflow-y-auto pr-1 custom-scrollbar">
              {visibleCustomers.length === 0 ? (
                <p className="text-xs text-slate-400 text-center py-8">
                  {csvRows.length === 0 ? 'Cargue un CSV para listar clientes' : `Sin datos para ${currentUser.country}`}
                </p>
              ) : (
                visibleCustomers.map(cust => (
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
                      <p className="text-xs font-extrabold text-sky-700">${cust.totalCost.toFixed(2)}</p>
                      <p className="text-[9px] text-slate-400">USD</p>
                    </div>
                  </div>
                ))
              )}
            </div>

            {/* PANEL DE ASIGNACIÓN DE ROLES Y REGISTRO DE USUARIOS (SOLO ADMINISTRADOR) */}
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

                {/* GESTIÓN RÁPIDA DE USUARIOS AUTORIZADOS CON BOTÓN ELIMINAR */}
                <div className="space-y-1.5 pt-2">
                  <p className="text-[10px] font-bold text-slate-400 uppercase">Usuarios Autorizados ({usersList.length})</p>
                  <div className="max-h-28 overflow-y-auto space-y-1 custom-scrollbar">
                    {usersList.map(user => (
                      <div key={user.id} className="flex items-center justify-between bg-slate-50 p-2 rounded-lg text-[11px]">
                        <div className="truncate max-w-[150px]">
                          <p className="font-bold text-slate-800 truncate">{user.email}</p>
                          <span className="text-[9px] text-slate-500">{user.role === 'admin' ? 'Acceso Global Admin' : `Auditor: ${user.country}`}</span>
                        </div>

                        <div className="flex items-center space-x-1">
                          <button
                            onClick={() => handleToggleUserRole(user.id)}
                            className="text-[10px] font-bold bg-slate-200 hover:bg-sky-100 text-slate-700 hover:text-sky-800 px-2 py-1 rounded transition"
                            title="Cambiar Rol"
                          >
                            {user.role === 'admin' ? 'Hacer Usuario' : 'Hacer Admin'}
                          </button>

                          <button
                            onClick={() => handleDeleteUser(user.id, user.email)}
                            className="text-red-500 hover:text-red-700 hover:bg-red-50 p-1 rounded transition"
                            title="Eliminar acceso de usuario"
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
                      <p className="text-2xl font-black text-sky-700">${selectedCustomer.totalCost.toFixed(2)} USD</p>
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
                      <span>Distribución por Categoría de Servicio (meterCategory)</span>
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
                        {selectedCustomer.records.map((r, i) => (
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
                            <td className="p-2.5 whitespace-nowrap">{r.date || r.billingPeriodEndDate}</td>
                            <td className="p-2.5">{r.unitOfMeasure}</td>
                            <td className="p-2.5">{r.quantity}</td>
                            <td className="p-2.5 font-extrabold text-emerald-600 whitespace-nowrap">
                              ${parseFloat(String(r[selectedCostKey] || r.paygCostInUsd || 0)).toFixed(2)}
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