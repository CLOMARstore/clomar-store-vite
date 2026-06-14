import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { supabase, hasSupabaseConfig } from './supabaseClient';
import * as XLSX from 'xlsx';
import { LogOut, Menu, Search, ShoppingCart, X } from 'lucide-react';
import './styles.css';

const money = (value) => `S/ ${Number(value || 0).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const todayISO = () => {
  const d = new Date();
  return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
};
const asNum = (v) => Number(v || 0);
const fmtDate = (value) => value ? new Date(value).toLocaleString('es-PE') : '';
const ticketText = (sale, items = [], payments = []) => {
  const lines = [];
  lines.push('CLOMAR STORE');
  lines.push('Comprobante interno');
  lines.push('------------------------------');
  lines.push(`Boleta: B${sale?.receipt_number || '—'}`);
  lines.push(`Fecha: ${fmtDate(sale?.created_at || new Date())}`);
  lines.push(`Cliente: ${sale?.customer_name || 'Cliente'}`);
  lines.push(`Pago: ${sale?.payment_method || 'Efectivo'}`);
  lines.push('------------------------------');
  items.forEach(it => lines.push(`${it.qty} x ${it.product_name || it.name}  ${money(asNum(it.subtotal) || asNum(it.qty)*asNum(it.price))}`));
  lines.push('------------------------------');
  lines.push(`TOTAL: ${money(sale?.total || 0)}`);
  if (payments.length) {
    lines.push('');
    lines.push('ABONOS:');
    payments.forEach(p => lines.push(`${fmtDate(p.created_at)} · ${p.method || p.payment_method}: ${money(p.amount)}`));
  }
  lines.push('Gracias por su compra.');
  return lines.join('\n');
};
const downloadText = (filename, content) => {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
};

const demoProducts = [
  { id: 'demo-1', code: '0001', name: 'Zapatillas deportivas Newton Nimble Leather', category: 'Calzado', price: 380, cost: 220, stock: 5, stock_min: 2, image_url: '', image_path: '', brand: '', size: '', color: '', description: '', barcode: '', active: true },
  { id: 'demo-2', code: '0002', name: 'Sombrero para el sol Bora Bora Booney', category: 'Accesorios', price: 80, cost: 40, stock: 2, stock_min: 2, image_url: '', image_path: '', brand: '', size: '', color: '', description: '', barcode: '', active: true },
  { id: 'demo-3', code: '0003', name: 'Camisa de popelina de manga larga para hombre', category: 'Ropa', price: 120, cost: 55, stock: 1, stock_min: 2, image_url: '', image_path: '', brand: '', size: '', color: '', description: '', barcode: '', active: true },
];

const DEFAULT_STORE_ID = '00000000-0000-0000-0000-000000000001';
const APP_ICON = '/logo-clomar-icon.png';
const APP_LOGO_FULL = '/logo-clomar-full.png';
const logoSrc = (store) => store?.logo_url || APP_ICON;
const productImageSrc = (product) => product?.image_url || APP_ICON;

const CODE128_PATTERNS = [
  '212222','222122','222221','121223','121322','131222','122213','122312','132212','221213','221312','231212','112232','122132','122231','113222','123122','123221','223211','221132','221231','213212','223112','312131','311222','321122','321221','312212','322112','322211','212123','212321','232121','111323','131123','131321','112313','132113','132311','211313','231113','231311','112133','112331','132131','113123','113321','133121','313121','211331','231131','213113','213311','213131','311123','311321','331121','312113','312311','332111','314111','221411','431111','111224','111422','121124','121421','141122','141221','112214','112412','122114','122411','142112','142211','241211','221114','413111','241112','134111','111242','121142','121241','114212','124112','124211','411212','421112','421211','212141','214121','412121','111143','111341','131141','114113','114311','411113','411311','113141','114131','311141','411131','211412','211214','211232','2331112'
];
const code128Sequence = (value = '') => {
  const text = String(value || '').trim();
  if (!text) return [104, 16, 106];
  const chars = Array.from(text).map((ch) => {
    const code = ch.charCodeAt(0);
    return code >= 32 && code <= 127 ? code - 32 : 16;
  });
  let checksum = 104;
  chars.forEach((code, idx) => { checksum += code * (idx + 1); });
  return [104, ...chars, checksum % 103, 106];
};
function BarcodeSVG({ value, height = 46 }) {
  const sequence = code128Sequence(value);
  let x = 0;
  const bars = [];
  sequence.forEach((code, codeIndex) => {
    const pattern = CODE128_PATTERNS[code] || CODE128_PATTERNS[16];
    let black = true;
    for (const widthChar of pattern) {
      const w = Number(widthChar || 1);
      if (black) bars.push(<rect key={`${codeIndex}-${x}`} x={x} y="0" width={w} height={height} />);
      x += w;
      black = !black;
    }
  });
  return <svg className="barcode-svg" viewBox={`0 0 ${x} ${height}`} preserveAspectRatio="none" role="img" aria-label={`Código de barras ${value}`}>{bars}</svg>;
}
const qrUrl = (value) => `https://api.qrserver.com/v1/create-qr-code/?size=180x180&margin=8&data=${encodeURIComponent(String(value || ''))}`;
const productScanCode = (product) => String(product?.barcode || product?.code || product?.id || '').trim();
const cleanFileName = (name = 'producto') => String(name).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z0-9._-]+/g, '-').replace(/-+/g, '-').toLowerCase();

const normalizeText = (value = '') => String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
const normalizeHeader = (value = '') => normalizeText(value).replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
const categoryPrefix = (categoryName = '') => {
  const n = normalizeText(categoryName);
  if (n.includes('ropa hombre')) return 'RH';
  if (n.includes('ropa mujer')) return 'RM';
  if (n.includes('calzado')) return 'CAL';
  if (n.includes('accesorios')) return 'ACC';
  if (n.includes('belleza') || n.includes('cuidado')) return 'BEL';
  if (n.includes('hogar')) return 'HOG';
  if (n.includes('fiesta') || n.includes('pinateria')) return 'FIE';
  if (n.includes('bazar')) return 'BAZ';
  return 'GEN';
};
const canonicalProductField = (key) => ({
  codigo: 'code', code: 'code', cod: 'code', codigo_interno: 'code', sku: 'code',
  barcode: 'barcode', codigo_barras: 'barcode', codigo_de_barras: 'barcode', barra: 'barcode', ean: 'barcode', upc: 'barcode',
  nombre: 'name', producto: 'name', name: 'name', descripcion_producto: 'name',
  categoria: 'category', category: 'category',
  subcategoria: 'subcategory', sub_category: 'subcategory', subcategory: 'subcategory',
  marca: 'brand', brand: 'brand',
  talla: 'size', size: 'size', medida: 'size',
  color: 'color',
  descripcion: 'description', description: 'description', detalles: 'description',
  costo: 'cost', cost: 'cost', costo_compra: 'cost', compra: 'cost',
  precio: 'price', price: 'price', precio_venta: 'price', venta: 'price',
  stock: 'stock', cantidad: 'stock', unidades: 'stock',
  stock_min: 'stock_min', stock_minimo: 'stock_min', minimo: 'stock_min',
  imagen_url: 'image_url', image_url: 'image_url', url_imagen: 'image_url', imagen: 'image_url',
  activo: 'active', active: 'active', estado: 'active',
}[key] || key);
const parseMoneyLike = (value) => {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/S\/?/gi, '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
  return Number(cleaned || 0);
};
const boolActive = (value) => {
  if (value === false) return false;
  const n = normalizeText(value);
  return !['no', 'false', 'inactivo', '0', 'desactivado'].includes(n);
};
const ROLE_LABELS = {
  dueno: 'Dueño',
  admin: 'Administrador',
  cajero: 'Cajero',
  almacen: 'Almacén',
  lectura: 'Solo lectura',
};
const ROLE_HOME = {
  dueno: 'panel',
  admin: 'panel',
  cajero: 'ventas',
  almacen: 'productos',
  lectura: 'reportes',
};
const MODULE_PERMISSIONS = {
  panel: ['dueno', 'admin', 'lectura'],
  ventas: ['dueno', 'admin', 'cajero'],
  creditos: ['dueno', 'admin', 'cajero'],
  caja: ['dueno', 'admin', 'cajero'],
  reportes: ['dueno', 'admin', 'lectura'],
  productos: ['dueno', 'admin', 'almacen'],
  categorias: ['dueno', 'admin', 'almacen'],
  etiquetas: ['dueno', 'admin', 'almacen'],
  inventario: ['dueno', 'admin', 'almacen', 'lectura'],
  ingreso: ['dueno', 'admin', 'almacen'],
  clientes: ['dueno', 'admin', 'cajero'],
  usuarios: ['dueno', 'admin'],
  tienda: ['dueno', 'admin'],
  herramientas: ['dueno'],
};
const canAccess = (role, moduleKey) => MODULE_PERMISSIONS[moduleKey]?.includes(role || 'cajero');
const firstAllowedModule = (role) => ROLE_HOME[role] || 'ventas';


function useAuth() {
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!hasSupabaseConfig) {
      setLoading(false);
      return;
    }
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  return { session, loading };
}


function useUserProfile(session) {
  const [profile, setProfile] = useState(null);
  const [store, setStore] = useState(null);
  const [loading, setLoading] = useState(true);

  async function loadProfile() {
    if (!hasSupabaseConfig || !session?.user?.id) {
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data: profileData, error: profileError } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', session.user.id)
      .single();

    const fallbackProfile = {
      id: session.user.id,
      store_id: DEFAULT_STORE_ID,
      email: session.user.email,
      full_name: session.user.email,
      role: 'cajero',
      status: 'Activo',
    };
    const nextProfile = profileError ? fallbackProfile : (profileData || fallbackProfile);
    setProfile(nextProfile);

    const { data: storeData } = await supabase
      .from('stores')
      .select('*')
      .eq('id', nextProfile.store_id || DEFAULT_STORE_ID)
      .single();
    setStore(storeData || { id: DEFAULT_STORE_ID, name: 'Clomar Store Pro' });
    setLoading(false);
  }

  useEffect(() => { loadProfile(); }, [session?.user?.id]);
  return { profile, store, loading, reload: loadProfile };
}

function roleMeta(profile) {
  return ROLE_LABELS[profile?.role] || 'Usuario';
}

function AccessDenied({ profile, setCurrent }) {
  return (
    <div className="page">
      <div className="hero compact-hero">
        <h1>🔒 Acceso restringido</h1>
        <p>Tu rol actual es {roleMeta(profile)}. No tienes permiso para este módulo.</p>
      </div>
      <section className="card compact-card">
        <p className="muted">Regresa a un módulo permitido o solicita al dueño cambiar tus permisos.</p>
        <button className="primary-btn" onClick={() => setCurrent(firstAllowedModule(profile?.role))}>Ir a mi panel</button>
      </section>
    </div>
  );
}

function InactiveUser({ profile }) {
  return (
    <main className="login-page">
      <section className="login-card">
        <div className="brand-logo"><img src={APP_ICON} alt="Clomar Store" /></div>
        <h1>Usuario inactivo</h1>
        <p>La cuenta {profile?.email || ''} está inactiva. Contacta al dueño o administrador.</p>
        <button className="primary-btn" onClick={() => supabase?.auth.signOut()}>Cerrar sesión</button>
      </section>
    </main>
  );
}

function Login() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  async function signIn(e) {
    e.preventDefault();
    setError('');
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { error: signError } = await supabase.auth.signInWithPassword({ email, password });
    if (signError) setError('Correo o contraseña incorrectos.');
    setLoading(false);
  }

  return (
    <main className="login-page">
      <section className="login-card">
        <div className="brand-logo"><img src={APP_ICON} alt="Clomar Store" /></div>
        <h1>Clomar Store</h1>
        <p>POS rápido para ventas, inventario, caja y créditos.</p>
        {!hasSupabaseConfig && (
          <div className="warning-box">
            Falta configurar Supabase. Copia <strong>.env.example</strong> como <strong>.env</strong> y coloca tus claves.
          </div>
        )}
        <form onSubmit={signIn}>
          <label>Correo del usuario</label>
          <input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="usuario@tienda.com" autoComplete="email" />
          <label>Contraseña</label>
          <input value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Tu contraseña" type="password" autoComplete="current-password" />
          {error && <div className="error-box">{error}</div>}
          <button disabled={loading || !hasSupabaseConfig} className="primary-btn">{loading ? 'Ingresando...' : 'Entrar'}</button>
        </form>
      </section>
    </main>
  );
}

function Sidebar({ current, setCurrent, open, setOpen, session, profile, store }) {
  const role = profile?.role || 'cajero';
  const sections = [
    { title: 'Gestionar negocio', items: [
      ['panel', '📊', 'Panel dueño'],
      ['ventas', '🧾', 'Ventas'],
      ['creditos', '💳', 'Créditos'],
      ['caja', '💰', 'Caja'],
      ['reportes', '📈', 'Reportes'],
    ]},
    { title: 'Productos e inventario', items: [
      ['productos', '📦', 'Productos'],
      ['categorias', '🏷️', 'Categorías'],
      ['etiquetas', '🏷️', 'Etiquetas'],
      ['inventario', '📘', 'Inventario'],
      ['ingreso', '📥', 'Ingreso mercadería'],
    ]},
    { title: 'Contactos', items: [['clientes', '👥', 'Clientes']]},
    { title: 'Administración', items: [
      ['usuarios', '🧑‍💼', 'Usuarios'],
      ['tienda', '🏪', 'Tienda'],
      ['herramientas', '🛠️', 'Herramientas'],
    ]},
  ].map(section => ({ ...section, items: section.items.filter(([key]) => canAccess(role, key)) })).filter(section => section.items.length);
  return (
    <aside className={`sidebar ${open ? 'open' : ''}`}>
      <div className="sidebar-head">
        <div className="mini-logo"><img src={logoSrc(store)} alt="Logo tienda" /></div>
        <div>
          <strong>{store?.name || 'Clomar Store Pro'}</strong>
          <small>{profile?.full_name || session?.user?.email || 'Usuario'} · {roleMeta(profile)}</small>
        </div>
        <button className="ghost mobile-only" onClick={() => setOpen(false)}><X size={18}/></button>
      </div>
      {sections.map((section) => (
        <div key={section.title} className="menu-section">
          <span>{section.title}</span>
          {section.items.map(([key, icon, label]) => (
            <button key={key} className={current === key ? 'active' : ''} onClick={() => { setCurrent(key); setOpen(false); }}>
              <span>{icon}</span>{label}
            </button>
          ))}
        </div>
      ))}
      <button className="logout" onClick={() => supabase?.auth.signOut()}><LogOut size={16}/> Cerrar sesión</button>
    </aside>
  );
}

function Header({ setOpen, current, profile, store }) {
  const titleMap = {
    panel: 'Panel dueño', ventas: 'Venta rápida', creditos: 'Créditos', caja: 'Caja diaria', reportes: 'Reportes', productos: 'Productos', categorias: 'Categorías', etiquetas: 'Etiquetas QR y barras', inventario: 'Inventario', ingreso: 'Ingreso de mercadería', clientes: 'Clientes', usuarios: 'Usuarios y roles', tienda: 'Configuración de tienda', herramientas: 'Herramientas'
  };
  return (
    <header className="app-header">
      <button className="ghost mobile-only" onClick={() => setOpen(true)}><Menu/></button>
      <div>
        <h2>{titleMap[current]}</h2>
        <p>{store?.name || 'Clomar Store Pro'} · {roleMeta(profile)}</p>
      </div>
    </header>
  );
}

function useProducts(profile) {
  const [products, setProducts] = useState([]);
  const [loading, setLoading] = useState(true);

  async function loadProducts() {
    if (!hasSupabaseConfig) {
      setProducts(demoProducts);
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data, error } = await supabase
      .from('products')
      .select('id,code,name,category,subcategory,category_id,subcategory_id,price,cost,stock,stock_min,status,store_id,image_url,image_path,brand,size,color,description,barcode,active')
      .eq('status', 'Activo')
      .eq('active', true)
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('name');
    if (!error) setProducts(data || []);
    setLoading(false);
  }

  useEffect(() => { loadProducts(); }, [profile?.store_id]);
  return { products, loading, reload: loadProducts };
}

function useCustomers(profile) {
  const [customers, setCustomers] = useState([]);
  const [loading, setLoading] = useState(true);

  async function loadCustomers() {
    if (!hasSupabaseConfig) {
      setCustomers([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data, error } = await supabase
      .from('customers')
      .select('*')
      .eq('status', 'Activo')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('name');
    if (!error) setCustomers(data || []);
    setLoading(false);
  }

  useEffect(() => { loadCustomers(); }, [profile?.store_id]);
  return { customers, loading, reload: loadCustomers };
}


function useCategories(profile) {
  const [categories, setCategories] = useState([]);
  const [loading, setLoading] = useState(true);

  async function loadCategories() {
    if (!hasSupabaseConfig) {
      setCategories([
        { id: 'cat-demo-1', name: 'Ropa hombre', parent_id: null, sort_order: 1, active: true },
        { id: 'cat-demo-2', name: 'Ropa mujer', parent_id: null, sort_order: 2, active: true },
        { id: 'cat-demo-3', name: 'Calzado', parent_id: null, sort_order: 3, active: true },
      ]);
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data, error } = await supabase
      .from('product_categories')
      .select('id,store_id,name,description,parent_id,sort_order,active')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .eq('active', true)
      .order('sort_order', { ascending: true })
      .order('name', { ascending: true });
    if (!error) setCategories(data || []);
    setLoading(false);
  }

  useEffect(() => { loadCategories(); }, [profile?.store_id]);
  return {
    categories: categories.filter(c => !c.parent_id),
    subcategories: categories.filter(c => c.parent_id),
    allCategories: categories,
    loading,
    reload: loadCategories,
  };
}

function useSales(profile) {
  const [sales, setSales] = useState([]);
  const [loading, setLoading] = useState(false);
  async function loadSales(limit = 50) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data, error } = await supabase
      .from('sales')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false })
      .limit(limit);
    if (!error) setSales(data || []);
    setLoading(false);
  }
  useEffect(() => { loadSales(); }, [profile?.store_id]);
  return { sales, loading, reload: loadSales };
}

function useCashMovements(profile) {
  const [movements, setMovements] = useState([]);
  const [loading, setLoading] = useState(false);
  async function loadMovements(limit = 80) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data, error } = await supabase
      .from('cash_movements')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false })
      .limit(limit);
    if (!error) setMovements(data || []);
    setLoading(false);
  }
  useEffect(() => { loadMovements(); }, [profile?.store_id]);
  return { movements, loading, reload: loadMovements };
}

function useStockMovements(profile) {
  const [movements, setMovements] = useState([]);
  const [loading, setLoading] = useState(false);
  async function loadMovements(limit = 80) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data, error } = await supabase
      .from('stock_movements')
      .select('*, products(name, code)')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false })
      .limit(limit);
    if (!error) setMovements(data || []);
    setLoading(false);
  }
  useEffect(() => { loadMovements(); }, [profile?.store_id]);
  return { movements, loading, reload: loadMovements };
}

function Kpi({ label, value, helper }) {
  return <div className="kpi"><span>{label}</span><strong>{value}</strong><small>{helper}</small></div>;
}

function Panel({ products, profile }) {
  const { sales } = useSales(profile);
  const { movements } = useCashMovements(profile);
  const stockCritico = products.filter(p => asNum(p.stock) <= asNum(p.stock_min ?? 2));
  const today = todayISO();
  const salesToday = sales.filter(s => String(s.created_at || '').slice(0,10) === today);
  const ingresosToday = movements.filter(m => String(m.created_at || '').slice(0,10) === today && ['Ingreso','Apertura'].includes(m.type));
  const egresosToday = movements.filter(m => String(m.created_at || '').slice(0,10) === today && ['Egreso','Compra','Retiro','Compra crédito'].includes(m.type));
  const totalVentas = salesToday.reduce((s, v) => s + asNum(v.total), 0);
  const cajaNeta = ingresosToday.reduce((s, v) => s + asNum(v.amount), 0) - egresosToday.reduce((s, v) => s + asNum(v.amount), 0);
  return (
    <div className="page compact-page">
      <div className="hero compact-hero">
        <h1>📊 Panel dueño</h1>
        <p>Resumen rápido del negocio. Carga ligera para evitar lag.</p>
      </div>
      <div className="kpi-grid">
        <Kpi label="Ventas hoy" value={money(totalVentas)} helper={`${salesToday.length} comprobantes`} />
        <Kpi label="Caja neta" value={money(cajaNeta)} helper="Ingresos menos egresos" />
        <Kpi label="Productos" value={products.length} helper="Activos" />
        <Kpi label="Stock crítico" value={stockCritico.length} helper="Revisar reposición" />
      </div>
      <div className="two-col">
        <section className="card compact-card">
          <h3>Últimas ventas</h3>
          {sales.slice(0, 6).length ? sales.slice(0, 6).map(s => (
            <div className="list-row" key={s.id}><span>B{s.receipt_number || '—'} · {s.customer_name || 'Cliente'}</span><strong>{money(s.total)}</strong></div>
          )) : <p className="muted">Todavía no hay ventas registradas.</p>}
        </section>
        <section className="card compact-card">
          <h3>Stock crítico</h3>
          {stockCritico.length ? stockCritico.slice(0, 8).map(p => (
            <div className="list-row" key={p.id}><span>{p.name}</span><strong>{asNum(p.stock)}</strong></div>
          )) : <p className="muted">No hay productos críticos.</p>}
        </section>
      </div>
    </div>
  );
}

function POS({ products, reloadProducts, customers, profile }) {
  const [query, setQuery] = useState('');
  const [cart, setCart] = useState([]);
  const [method, setMethod] = useState('Efectivo');
  const [customer, setCustomer] = useState('Cliente');
  const [saving, setSaving] = useState(false);
  const [lastTicket, setLastTicket] = useState(null);
  const [scanOpen, setScanOpen] = useState(false);
  const [scanStatus, setScanStatus] = useState('');
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const normalized = query.trim().toLowerCase();
  const activeProducts = useMemo(() => products.filter(p => p.active !== false), [products]);
  const matches = useMemo(() => {
    const base = !normalized ? activeProducts.slice(0, 12) : activeProducts.filter(p => `${p.code} ${p.barcode || ''} ${p.name} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''} ${p.color || ''} ${p.size || ''}`.toLowerCase().includes(normalized)).slice(0, 20);
    return base;
  }, [activeProducts, normalized]);
  const total = cart.reduce((sum, item) => sum + asNum(item.price) * asNum(item.qty), 0);

  function findProductByBarcode(value) {
    const clean = String(value || '').trim().toLowerCase();
    if (!clean) return null;
    return activeProducts.find(p => String(p.barcode || '').trim().toLowerCase() === clean || String(p.code || '').trim().toLowerCase() === clean) || null;
  }

  function addProduct(product) {
    if (asNum(product.stock) <= 0) return alert('Producto sin stock disponible.');
    setLastTicket(null);
    setCart(prev => {
      const found = prev.find(x => x.id === product.id);
      if (found) {
        if (asNum(found.qty) + 1 > asNum(product.stock)) return prev;
        return prev.map(x => x.id === product.id ? { ...x, qty: asNum(x.qty) + 1 } : x);
      }
      return [...prev, { ...product, qty: 1 }];
    });
  }

  function processBarcode(value, source = 'manual') {
    const clean = String(value || '').trim();
    if (!clean) return;
    const product = findProductByBarcode(clean);
    if (product) {
      addProduct(product);
      setQuery('');
      setScanStatus(`Agregado al carrito: ${product.name}`);
      if (source === 'camera') setScanOpen(false);
      return;
    }
    setQuery(clean);
    setScanStatus(`Código no encontrado: ${clean}`);
    if (source !== 'camera') alert(`Código no encontrado: ${clean}`);
  }

  function handleSearchKeyDown(e) {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const clean = query.trim();
    if (!clean) return;
    const exact = findProductByBarcode(clean);
    if (exact) { processBarcode(clean, 'lector'); return; }
    if (matches.length === 1) {
      addProduct(matches[0]);
      setQuery('');
      setScanStatus(`Agregado al carrito: ${matches[0].name}`);
      return;
    }
    alert('No se encontró un producto exacto. Escanea el código de barras o busca por nombre.');
  }

  function stopScanner() {
    if (streamRef.current) {
      streamRef.current.getTracks().forEach(track => track.stop());
      streamRef.current = null;
    }
    setScanOpen(false);
  }

  useEffect(() => {
    if (!scanOpen) return;
    let cancelled = false;
    let raf = 0;

    async function startScanner() {
      try {
        if (!navigator.mediaDevices?.getUserMedia) {
          setScanStatus('Este navegador no permite usar la cámara. Usa lector físico o escribe el código.');
          return;
        }
        if (!('BarcodeDetector' in window)) {
          setScanStatus('Tu navegador no tiene lector de código por cámara. Usa Chrome/Android, lector físico USB/Bluetooth o escribe el código manualmente.');
          return;
        }

        setScanStatus('Solicitando permiso de cámara...');
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } }, audio: false });
        if (cancelled) { stream.getTracks().forEach(track => track.stop()); return; }
        streamRef.current = stream;
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          await videoRef.current.play();
        }

        const supportedFormats = window.BarcodeDetector.getSupportedFormats ? await window.BarcodeDetector.getSupportedFormats() : [];
        const preferred = ['ean_13','ean_8','code_128','code_39','upc_a','upc_e','itf','qr_code'];
        const formats = supportedFormats.length ? preferred.filter(f => supportedFormats.includes(f)) : preferred;
        const detector = new window.BarcodeDetector({ formats: formats.length ? formats : undefined });
        setScanStatus('Apunta la cámara al código de barras.');

        async function loop() {
          if (cancelled || !videoRef.current) return;
          try {
            const codes = await detector.detect(videoRef.current);
            if (codes?.length) {
              const raw = codes[0].rawValue || codes[0].rawData;
              if (raw) { processBarcode(raw, 'camera'); return; }
            }
          } catch (err) {
            // Continuar intentando mientras la cámara esté abierta.
          }
          raf = requestAnimationFrame(loop);
        }
        loop();
      } catch (err) {
        setScanStatus(`No se pudo abrir la cámara: ${err.message || err}`);
      }
    }

    startScanner();
    return () => {
      cancelled = true;
      if (raf) cancelAnimationFrame(raf);
      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
        streamRef.current = null;
      }
    };
  }, [scanOpen]);

  function updateQty(id, qty) {
    setCart(prev => prev.map(x => {
      if (x.id !== id) return x;
      const next = Math.max(1, Math.min(asNum(x.stock), asNum(qty || 1)));
      return { ...x, qty: next };
    }));
  }
  function removeItem(id) { setCart(prev => prev.filter(x => x.id !== id)); }

  async function checkout() {
    if (!cart.length || saving) return;
    if (!hasSupabaseConfig) { alert('Venta demo registrada. Configura Supabase para guardar.'); setCart([]); return; }
    setSaving(true);
    const meta = { store_id: profile?.store_id || DEFAULT_STORE_ID, user_id: profile?.id || null };
    const salePayload = { customer_name: customer || 'Cliente', payment_method: method, total, status: method === 'Crédito' ? 'Crédito' : 'Pagado', ...meta };
    const { data: sale, error } = await supabase.from('sales').insert(salePayload).select().single();
    if (error) { alert(error.message); setSaving(false); return; }
    const items = cart.map(item => ({ sale_id: sale.id, product_id: item.id, product_name: item.name, qty: item.qty, price: item.price, subtotal: asNum(item.qty) * asNum(item.price), store_id: profile?.store_id || DEFAULT_STORE_ID }));
    await supabase.from('sale_items').insert(items);
    for (const item of cart) {
      await supabase.from('products').update({ stock: asNum(item.stock) - asNum(item.qty) }).eq('id', item.id);
      await supabase.from('stock_movements').insert({ product_id: item.id, type: 'Salida', qty: item.qty, note: `Venta B${sale.receipt_number || sale.id}`, ...meta });
    }
    await supabase.from('cash_movements').insert({ type: method === 'Crédito' ? 'Crédito' : 'Ingreso', payment_method: method, amount: total, note: `Venta B${sale.receipt_number || sale.id}`, ...meta });
    setLastTicket({ sale, items });
    setCart([]); setSaving(false); reloadProducts();
  }

  return (
    <div className="page pos-page">
      <div className="hero compact-hero"><h1>🧾 Venta rápida</h1><p>Busca por nombre, escanea código con lector físico o usa la cámara del celular.</p></div>
      <div className="pos-layout">
        <section className="card compact-card">
          <div className="barcode-tools">
            <div className="search-box barcode-search"><Search size={18}/><input value={query} onChange={(e)=>setQuery(e.target.value)} onKeyDown={handleSearchKeyDown} placeholder="Buscar o escanear código de barras..." autoFocus /></div>
            <button className="secondary-btn scan-btn" type="button" onClick={()=>setScanOpen(true)}>📷 Escanear con celular</button>
          </div>
          <div className="scanner-help">Lector físico: enfoca el buscador y escanea. Cámara: abre el escáner y apunta al código.</div>
          {scanStatus && <div className="scan-status">{scanStatus}</div>}
          {scanOpen && (
            <div className="scanner-panel">
              <div className="scanner-head"><strong>Escáner con cámara</strong><button className="icon-btn" type="button" onClick={stopScanner}>×</button></div>
              <div className="scanner-frame"><video ref={videoRef} muted playsInline /></div>
              <p className="muted">Usa la cámara trasera del celular. Si no detecta, escribe el código manualmente en el buscador.</p>
            </div>
          )}
          <div className="product-list">
            {matches.map(product => (
              <button key={product.id} className="product-row product-row-media" onClick={() => addProduct(product)}>
                <img className="product-thumb" src={productImageSrc(product)} alt={product.name} />
                <div className="product-row-info"><strong>{product.name}</strong><small>{product.code} · {product.category}{product.subcategory ? ` / ${product.subcategory}` : ''} · {product.brand || 'Sin marca'} · Stock {asNum(product.stock)}</small></div>
                <b>{money(product.price)}</b>
              </button>
            ))}
            {!matches.length && <p className="muted">No se encontraron productos.</p>}
          </div>
        </section>
        <aside className="card compact-card cart-card">
          <h3><ShoppingCart size={20}/> Carrito</h3>
          {cart.length === 0 ? <p className="muted">Agrega productos para vender.</p> : cart.map(item => (
            <div className="cart-row" key={item.id}>
              <div><strong>{item.name}</strong><small>{money(item.price)} c/u · Stock {asNum(item.stock)}</small></div>
              <input type="number" value={item.qty} min="1" max={asNum(item.stock)} onChange={(e)=>updateQty(item.id, e.target.value)} />
              <strong>{money(asNum(item.qty)*asNum(item.price))}</strong>
              <button className="icon-btn" onClick={()=>removeItem(item.id)}>×</button>
            </div>
          ))}
          <div className="cart-total"><span>Total</span><strong>{money(total)}</strong></div>
          <select value={customer} onChange={(e)=>setCustomer(e.target.value)}>
            <option>Cliente</option>
            {customers.map(c => <option key={c.id} value={c.name}>{c.name}</option>)}
          </select>
          <select value={method} onChange={(e)=>setMethod(e.target.value)}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option><option>Crédito</option></select>
          <button className="primary-btn" disabled={!cart.length || saving} onClick={checkout}>{saving ? 'Guardando...' : method === 'Crédito' ? 'Registrar crédito' : 'Cobrar'}</button>
          {lastTicket && (
            <div className="ticket-box">
              <h4>✅ Venta registrada</h4>
              <p><strong>Boleta interna:</strong> B{lastTicket.sale.receipt_number}</p>
              <p><strong>Total:</strong> {money(lastTicket.sale.total)}</p>
              <div className="ticket-actions">
                <button className="secondary-btn" onClick={() => downloadText(`boleta-B${lastTicket.sale.receipt_number}.txt`, ticketText(lastTicket.sale, lastTicket.items))}>Descargar ticket</button>
                <button className="secondary-btn" onClick={() => window.print()}>Imprimir pantalla</button>
              </div>
            </div>
          )}
        </aside>
      </div>
    </div>
  );
}



function CategoriesAdmin({ profile, categories = [], subcategories = [], products = [], reloadCategories }) {
  const [form, setForm] = useState({ name: '', type: 'principal', parent_id: '', description: '', sort_order: '100' });
  const [saving, setSaving] = useState(false);
  const [query, setQuery] = useState('');
  const [expanded, setExpanded] = useState({});

  const productCountFor = (cat) => products.filter(p => p.active !== false && (p.category_id === cat.id || p.subcategory_id === cat.id || p.category === cat.name || p.subcategory === cat.name)).length;
  const visibleCategories = categories.filter(cat => {
    const children = subcategories.filter(s => s.parent_id === cat.id);
    const text = `${cat.name} ${children.map(c => c.name).join(' ')}`.toLowerCase();
    return text.includes(query.toLowerCase());
  });

  function resetForm() {
    setForm({ name: '', type: 'principal', parent_id: '', description: '', sort_order: '100' });
  }

  async function save(e) {
    e.preventDefault();
    if (!form.name.trim()) return alert('Coloca el nombre de la categoría o subcategoría.');
    if (form.type === 'subcategoria' && !form.parent_id) return alert('Selecciona la categoría principal de esta subcategoría.');
    setSaving(true);
    const payload = {
      store_id: profile?.store_id || DEFAULT_STORE_ID,
      name: form.name.trim(),
      parent_id: form.type === 'subcategoria' ? form.parent_id : null,
      description: form.description.trim(),
      sort_order: asNum(form.sort_order || 100),
      active: true,
    };
    const { error } = await supabase.from('product_categories').insert(payload);
    setSaving(false);
    if (error) return alert(error.message || 'No se pudo guardar la categoría.');
    resetForm();
    reloadCategories?.();
  }

  async function renameCategory(cat) {
    const nextName = prompt('Nuevo nombre de categoría/subcategoría:', cat.name);
    if (!nextName || !nextName.trim() || nextName.trim() === cat.name) return;
    const { error } = await supabase.from('product_categories').update({ name: nextName.trim() }).eq('id', cat.id);
    if (error) alert(error.message); else reloadCategories?.();
  }

  async function deactivateCategory(cat) {
    if (!confirm(`¿Desactivar ${cat.name}? No se borrarán los productos ya registrados.`)) return;
    const { error } = await supabase.from('product_categories').update({ active: false }).eq('id', cat.id);
    if (error) alert(error.message); else reloadCategories?.();
  }

  function toggle(id) {
    setExpanded(prev => ({ ...prev, [id]: !prev[id] }));
  }

  return (
    <div className="page">
      <div className="hero compact-hero"><h1>🏷️ Categorías</h1><p>Panel profesional para organizar productos por categoría principal y subcategoría.</p></div>
      <div className="category-toolbar card compact-card">
        <div>
          <h3>Resumen de categorías</h3>
          <p className="muted">{categories.length} categorías principales · {subcategories.length} subcategorías activas · {products.filter(p=>p.active!==false).length} productos activos</p>
        </div>
        <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar categoría o subcategoría..." /></div>
      </div>
      <div className="two-col category-admin-layout polished-categories">
        <form className="card form-grid category-form-card" onSubmit={save}>
          <h3>Nueva categoría</h3>
          <label>Tipo de registro
            <select value={form.type} onChange={e=>setForm({...form,type:e.target.value,parent_id:''})}>
              <option value="principal">Categoría principal</option>
              <option value="subcategoria">Subcategoría</option>
            </select>
          </label>
          {form.type === 'subcategoria' && (
            <label>Categoría padre
              <select value={form.parent_id} onChange={e=>setForm({...form,parent_id:e.target.value})}>
                <option value="">Selecciona categoría principal</option>
                {categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </label>
          )}
          <label>{form.type === 'principal' ? 'Nombre de categoría principal' : 'Nombre de subcategoría'}
            <input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder={form.type === 'principal' ? 'Ejemplo: Tecnología' : 'Ejemplo: Correas hombre'} />
          </label>
          <label>Descripción interna
            <input value={form.description} onChange={e=>setForm({...form,description:e.target.value})} placeholder="Opcional para control interno" />
          </label>
          <label>Orden
            <input value={form.sort_order} onChange={e=>setForm({...form,sort_order:e.target.value})} inputMode="numeric" />
          </label>
          <div className="form-actions-row">
            <button type="button" className="secondary-btn" onClick={resetForm}>Limpiar</button>
            <button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar'}</button>
          </div>
        </form>
        <section className="card compact-card category-panel-card">
          <div className="category-panel-head">
            <div>
              <h3>Mapa de categorías</h3>
              <p className="muted">Haz clic en “Ver” para desplegar subcategorías.</p>
            </div>
          </div>
          <div className="category-card-grid">
            {visibleCategories.map(cat => {
              const children = subcategories.filter(s => s.parent_id === cat.id);
              const catProducts = productCountFor(cat);
              const isOpen = expanded[cat.id] || query.trim();
              return (
                <article className="category-card-pro" key={cat.id}>
                  <div className="category-card-main">
                    <div className="category-icon-badge">🏷️</div>
                    <div>
                      <strong>{cat.name}</strong>
                      <small>{children.length} subcategorías · {catProducts} productos</small>
                    </div>
                  </div>
                  <div className="category-card-actions">
                    <button type="button" className="secondary-btn" onClick={()=>toggle(cat.id)}>{isOpen ? 'Ocultar' : 'Ver'}</button>
                    <button type="button" className="secondary-btn" onClick={()=>renameCategory(cat)}>Editar</button>
                    <button type="button" className="danger-mini-btn" onClick={()=>deactivateCategory(cat)}>Desactivar</button>
                  </div>
                  {isOpen && (
                    <div className="subcategory-list-pro">
                      {children.length ? children.map(sub => (
                        <div className="subcategory-row-pro" key={sub.id}>
                          <span>{sub.name}</span>
                          <small>{productCountFor(sub)} productos</small>
                          <button type="button" onClick={()=>renameCategory(sub)}>Editar</button>
                          <button type="button" className="danger-text-btn" onClick={()=>deactivateCategory(sub)}>Desactivar</button>
                        </div>
                      )) : <p className="muted empty-subcategory">Sin subcategorías registradas.</p>}
                    </div>
                  )}
                </article>
              );
            })}
            {!visibleCategories.length && <p className="muted">No se encontraron categorías con ese criterio.</p>}
          </div>
        </section>
      </div>
    </div>
  );
}

function Products({ products, reload, profile, categories = [], subcategories = [], reloadCategories }) {
  const emptyForm = { code:'', barcode:'', name:'', category_id:'', subcategory_id:'', category:'', subcategory:'', brand:'', size:'', color:'', description:'', price:'', cost:'', stock:'0', stock_min:'2', image_url:'' };
  const [form, setForm] = useState(emptyForm);
  const [imageFile, setImageFile] = useState(null);
  const [previewUrl, setPreviewUrl] = useState('');
  const [saving, setSaving] = useState(false);
  const [query, setQuery] = useState('');
  const filteredProducts = products.filter(p => `${p.code} ${p.barcode || ''} ${p.name} ${p.category || ''} ${p.subcategory || ''} ${p.brand || ''} ${p.color || ''} ${p.size || ''}`.toLowerCase().includes(query.toLowerCase()));
  const selectedSubcategories = subcategories.filter(c => c.parent_id === form.category_id);
  function setCategoryById(categoryId) {
    const cat = categories.find(c => c.id === categoryId);
    setForm({ ...form, category_id: categoryId, category: cat?.name || '', subcategory_id: '', subcategory: '' });
  }
  function setSubcategoryById(subcategoryId) {
    const sub = subcategories.find(c => c.id === subcategoryId);
    setForm({ ...form, subcategory_id: subcategoryId, subcategory: sub?.name || '' });
  }

  useEffect(() => {
    if (!imageFile) {
      setPreviewUrl(form.image_url || '');
      return;
    }
    const localUrl = URL.createObjectURL(imageFile);
    setPreviewUrl(localUrl);
    return () => URL.revokeObjectURL(localUrl);
  }, [imageFile, form.image_url]);

  async function uploadProductImage() {
    if (!imageFile) return { image_url: form.image_url || '', image_path: '' };
    const storeId = profile?.store_id || DEFAULT_STORE_ID;
    const filePath = `${storeId}/${Date.now()}-${cleanFileName(imageFile.name)}`;
    const { error: uploadError } = await supabase.storage
      .from('product-images')
      .upload(filePath, imageFile, { cacheControl: '3600', upsert: false });
    if (uploadError) throw uploadError;
    const { data } = supabase.storage.from('product-images').getPublicUrl(filePath);
    return { image_url: data?.publicUrl || '', image_path: filePath };
  }

  async function save(e) {
    e.preventDefault();
    if (!hasSupabaseConfig) return alert('Configura Supabase para guardar productos.');
    if (!form.name.trim()) return alert('Coloca el nombre del producto.');
    if (asNum(form.price) <= 0) return alert('Coloca un precio de venta válido.');
    setSaving(true);
    try {
      const imageData = await uploadProductImage();
      const payload = {
        code: form.code.trim(),
        barcode: form.barcode.trim(),
        name: form.name.trim(),
        category: form.category.trim() || 'General',
        subcategory: form.subcategory.trim(),
        category_id: form.category_id || null,
        subcategory_id: form.subcategory_id || null,
        brand: form.brand.trim(),
        size: form.size.trim(),
        color: form.color.trim(),
        description: form.description.trim(),
        price: asNum(form.price),
        cost: asNum(form.cost),
        stock: asNum(form.stock),
        stock_min: asNum(form.stock_min),
        status: 'Activo',
        active: true,
        image_url: imageData.image_url,
        image_path: imageData.image_path,
        store_id: profile?.store_id || DEFAULT_STORE_ID,
        created_by: profile?.id || null,
      };
      const { error } = await supabase.from('products').insert(payload);
      if (error) throw error;
      setForm(emptyForm);
      setImageFile(null);
      setPreviewUrl('');
      reload();
      alert('Producto guardado correctamente.');
    } catch (error) {
      alert(error.message || 'No se pudo guardar el producto.');
    } finally {
      setSaving(false);
    }
  }

  async function deactivateProduct(product) {
    if (!confirm(`¿Desactivar ${product.name}? No se borrará el historial de ventas.`)) return;
    const { error } = await supabase.from('products').update({ active: false, status: 'Inactivo' }).eq('id', product.id);
    if (error) alert(error.message); else reload();
  }

  return (
    <div className="page">
      <div className="hero compact-hero"><h1>📦 Productos con imágenes</h1><p>Crea artículos visuales con marca, talla, color, código de barras y foto.</p></div>
      <div className="two-col product-admin-layout">
        <form className="card form-grid" onSubmit={save}>
          <h3>Nuevo producto</h3>
          <div className="image-uploader">
            <div className="image-preview-box">
              <img src={previewUrl || APP_ICON} alt="Vista previa" />
              <small>{previewUrl ? 'Vista previa de imagen' : 'Sin imagen cargada'}</small>
            </div>
            <label>Subir imagen del producto<input type="file" accept="image/png,image/jpeg,image/webp" onChange={e=>setImageFile(e.target.files?.[0] || null)} /></label>
            <label>O pegar URL pública de imagen<input value={form.image_url} onChange={e=>{ setImageFile(null); setForm({...form,image_url:e.target.value}); }} placeholder="https://..." /></label>
          </div>
          <div className="form-split">
            <label>Código interno<input value={form.code} onChange={e=>setForm({...form,code:e.target.value})} placeholder="0004" /></label>
            <label>Código de barras<input value={form.barcode} onChange={e=>setForm({...form,barcode:e.target.value})} placeholder="Escanea o escribe el código" /></label>
          </div>
          <label>Nombre del producto<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder="Nombre comercial del producto" /></label>
          <div className="form-split">
            <label>Categoría
              <select value={form.category_id} onChange={e=>setCategoryById(e.target.value)}>
                <option value="">Selecciona categoría</option>
                {categories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </label>
            <label>Subcategoría
              <select value={form.subcategory_id} onChange={e=>setSubcategoryById(e.target.value)} disabled={!form.category_id}>
                <option value="">Selecciona subcategoría</option>
                {selectedSubcategories.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </label>
          </div>
          <label>Marca<input value={form.brand} onChange={e=>setForm({...form,brand:e.target.value})} placeholder="Marca" /></label>
          <div className="form-split">
            <label>Talla<input value={form.size} onChange={e=>setForm({...form,size:e.target.value})} placeholder="S, M, L, 38, 40" /></label>
            <label>Color<input value={form.color} onChange={e=>setForm({...form,color:e.target.value})} placeholder="Negro, rosa, azul" /></label>
          </div>
          <label>Descripción<input value={form.description} onChange={e=>setForm({...form,description:e.target.value})} placeholder="Material, modelo, detalles para catálogo" /></label>
          <div className="form-split">
            <label>Precio de venta<input value={form.price} inputMode="decimal" onChange={e=>setForm({...form,price:e.target.value})} placeholder="0.00" /></label>
            <label>Costo de compra<input value={form.cost} inputMode="decimal" onChange={e=>setForm({...form,cost:e.target.value})} placeholder="0.00" /></label>
          </div>
          <div className="form-split">
            <label>Stock inicial<input value={form.stock} inputMode="decimal" onChange={e=>setForm({...form,stock:e.target.value})} /></label>
            <label>Stock mínimo<input value={form.stock_min} inputMode="decimal" onChange={e=>setForm({...form,stock_min:e.target.value})} /></label>
          </div>
          <button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar producto'}</button>
        </form>
        <section className="card compact-card">
          <h3>Lista de productos</h3>
          <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar por nombre, marca, código o barcode..." /></div>
          <div className="product-card-list">
            {filteredProducts.map(p=>(
              <div className="product-card-item" key={p.id}>
                <img src={productImageSrc(p)} alt={p.name} />
                <div>
                  <strong>{p.name}</strong>
                  <small>{p.code || 'Sin código'} · {p.barcode || 'Sin barcode'} · {p.category || 'General'}{p.subcategory ? ` / ${p.subcategory}` : ''}</small>
                  <small>{p.brand || 'Sin marca'} · {p.size || 'Sin talla'} · {p.color || 'Sin color'}</small>
                  {p.description && <small>{p.description}</small>}
                </div>
                <div className="product-card-actions">
                  <b>{money(p.price)}</b>
                  <small>Stock {asNum(p.stock)}</small>
                  <button type="button" className="secondary-btn" onClick={()=>deactivateProduct(p)}>Desactivar</button>
                </div>
              </div>
            ))}
            {!filteredProducts.length && <p className="muted">No hay productos activos.</p>}
          </div>
        </section>
      </div>
    </div>
  );
}

function Customers({ customers, reload, profile }) {
  const [form, setForm] = useState({ name:'', phone:'', document:'', address:'', credit_limit:'0' });
  const [query, setQuery] = useState('');
  const filtered = customers.filter(c => `${c.name} ${c.phone} ${c.document}`.toLowerCase().includes(query.toLowerCase()));
  async function save(e) {
    e.preventDefault();
    if (!form.name.trim()) return alert('Coloca el nombre del cliente.');
    const payload = { ...form, credit_limit: asNum(form.credit_limit), status: 'Activo', store_id: profile?.store_id || DEFAULT_STORE_ID, created_by: profile?.id || null };
    const { error } = await supabase.from('customers').insert(payload);
    if (error) alert(error.message); else { setForm({ name:'', phone:'', document:'', address:'', credit_limit:'0' }); reload(); }
  }
  return (
    <div className="page">
      <div className="hero compact-hero"><h1>👥 Clientes</h1><p>Registra contactos para ventas, créditos y seguimiento.</p></div>
      <div className="two-col">
        <form className="card form-grid" onSubmit={save}>
          <label>Nombre<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder="Nombre del cliente" /></label>
          <label>Teléfono<input value={form.phone} onChange={e=>setForm({...form,phone:e.target.value})} placeholder="Celular o WhatsApp" /></label>
          <label>Documento<input value={form.document} onChange={e=>setForm({...form,document:e.target.value})} placeholder="DNI / RUC / documento" /></label>
          <label>Dirección<input value={form.address} onChange={e=>setForm({...form,address:e.target.value})} placeholder="Dirección o referencia" /></label>
          <label>Límite de crédito<input value={form.credit_limit} onChange={e=>setForm({...form,credit_limit:e.target.value})} inputMode="decimal" placeholder="0.00" /></label>
          <button className="primary-btn">Guardar cliente</button>
        </form>
        <section className="card compact-card">
          <h3>Lista de clientes</h3>
          <div className="search-box"><Search size={18}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Buscar cliente..." /></div>
          {filtered.map(c => <div className="list-row" key={c.id}><span>{c.name}<small>{c.phone || 'Sin teléfono'} · {c.document || 'Sin documento'}</small></span><strong>{money(c.credit_limit)}</strong></div>)}
          {!filtered.length && <p className="muted">No hay clientes registrados.</p>}
        </section>
      </div>
    </div>
  );
}

function Inventory({ products }) {
  const byCat = products.reduce((acc, p) => { (acc[p.category || 'General'] ||= []).push(p); return acc; }, {});
  return <div className="page"><div className="hero compact-hero"><h1>📘 Inventario</h1><p>Vista compacta por categoría y stock.</p></div>{Object.entries(byCat).map(([cat, items]) => <section className="card compact-card inventory-block" key={cat}><h3>{cat}</h3>{items.map(p=><div className="list-row inventory-product-row" key={p.id}><img className="product-thumb small" src={productImageSrc(p)} alt={p.name}/><span>{p.code} · {p.name}<small>{p.brand || 'Sin marca'} · {p.color || 'Sin color'} · Stock mínimo {asNum(p.stock_min)}</small></span><strong className={asNum(p.stock) <= asNum(p.stock_min) ? 'danger-text' : ''}>Stock {asNum(p.stock)}</strong></div>)}</section>)}</div>;
}

function StockEntry({ products, reloadProducts, profile }) {
  const [form, setForm] = useState({ product_id:'', provider:'', qty:'1', cost:'0', method:'Efectivo', paid:'0', note:'' });
  const selected = products.find(p => p.id === form.product_id) || products[0];
  const total = asNum(form.qty) * asNum(form.cost);
  useEffect(() => {
    if (!form.product_id && products[0]) setForm(f => ({ ...f, product_id: products[0].id, cost: String(products[0].cost || 0) }));
  }, [products]);
  useEffect(() => {
    if (form.method !== 'Crédito') setForm(f => ({ ...f, paid: String(total || 0) }));
  }, [form.qty, form.cost, form.method]);
  async function save(e) {
    e.preventDefault();
    if (!selected) return alert('Selecciona un producto.');
    if (asNum(form.qty) <= 0) return alert('La cantidad debe ser mayor a cero.');
    const newStock = asNum(selected.stock) + asNum(form.qty);
    const { error: prodError } = await supabase.from('products').update({ stock: newStock, cost: asNum(form.cost) }).eq('id', selected.id);
    if (prodError) return alert(prodError.message);
    const note = `Proveedor: ${form.provider || 'Sin proveedor'} · Método: ${form.method} · ${form.note || ''}`;
    await supabase.from('stock_movements').insert({ product_id: selected.id, type: 'Entrada', qty: asNum(form.qty), note, store_id: profile?.store_id || DEFAULT_STORE_ID, user_id: profile?.id || null });
    await supabase.from('cash_movements').insert({ type: form.method === 'Crédito' ? 'Compra crédito' : 'Compra', payment_method: form.method, amount: asNum(form.paid || total), note: `Ingreso mercadería: ${selected.name}. ${note}`, store_id: profile?.store_id || DEFAULT_STORE_ID, user_id: profile?.id || null });
    alert(`Ingreso registrado. Nuevo stock: ${newStock}`);
    setForm({ product_id: selected.id, provider:'', qty:'1', cost:String(form.cost || selected.cost || 0), method:'Efectivo', paid:'0', note:'' });
    reloadProducts();
  }
  return (
    <div className="page">
      <div className="hero compact-hero"><h1>📥 Ingreso de mercadería</h1><p>Repone stock, registra proveedor, costo y movimiento de caja.</p></div>
      <div className="two-col">
        <form className="card form-grid" onSubmit={save}>
          <label>Producto<select value={form.product_id} onChange={e=>{const p=products.find(x=>x.id===e.target.value); setForm({...form,product_id:e.target.value,cost:String(p?.cost || 0)})}}>{products.map(p=><option key={p.id} value={p.id}>{p.code} · {p.name} · Stock {asNum(p.stock)}</option>)}</select></label>
          <label>Proveedor<input value={form.provider} onChange={e=>setForm({...form,provider:e.target.value})} placeholder="Nombre del proveedor" /></label>
          <label>Cantidad ingresada<input value={form.qty} onChange={e=>setForm({...form,qty:e.target.value})} inputMode="decimal" /></label>
          <label>Costo unitario<input value={form.cost} onChange={e=>setForm({...form,cost:e.target.value})} inputMode="decimal" /></label>
          <label>Método de pago<select value={form.method} onChange={e=>setForm({...form,method:e.target.value})}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option><option>Crédito</option></select></label>
          {form.method === 'Crédito' && <label>Monto pagado / adelanto<input value={form.paid} onChange={e=>setForm({...form,paid:e.target.value})} inputMode="decimal" /></label>}
          <label>Observación<input value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Factura, guía, nota de compra..." /></label>
          <div className="total-box"><span>Total ingreso</span><strong>{money(total)}</strong><small>Se actualizará el stock al guardar.</small></div>
          <button className="primary-btn">Registrar ingreso de mercadería</button>
        </form>
        <section className="card compact-card">
          <h3>Producto seleccionado</h3>
          {selected ? <>
            <div className="list-row inventory-product-row"><img className="product-thumb small" src={productImageSrc(selected)} alt={selected.name}/><span>{selected.name}<small>{selected.code} · {selected.category}{selected.subcategory ? ` / ${selected.subcategory}` : ''} · {selected.brand || 'Sin marca'}</small></span><strong>{money(selected.price)}</strong></div>
            <div className="list-row"><span>Stock actual</span><strong>{asNum(selected.stock)}</strong></div>
            <div className="list-row"><span>Stock después</span><strong>{asNum(selected.stock) + asNum(form.qty)}</strong></div>
          </> : <p className="muted">No hay productos.</p>}
        </section>
      </div>
    </div>
  );
}

function CashPage({ profile }) {
  const { movements, reload } = useCashMovements(profile);
  const [form, setForm] = useState({ type:'Ingreso', method:'Efectivo', amount:'0', note:'' });
  const today = todayISO();
  const todayMovs = movements.filter(m => String(m.created_at || '').slice(0,10) === today);
  const ingresos = todayMovs.filter(m => ['Ingreso','Apertura'].includes(m.type)).reduce((s,m)=>s+asNum(m.amount),0);
  const egresosList = todayMovs.filter(m => ['Egreso','Compra','Retiro','Compra crédito'].includes(m.type));
  const egresos = egresosList.reduce((s,m)=>s+asNum(m.amount),0);
  const creditos = todayMovs.filter(m => String(m.type).includes('Crédito')).reduce((s,m)=>s+asNum(m.amount),0);
  const abonos = todayMovs.filter(m => String(m.note || '').toLowerCase().includes('abono')).reduce((s,m)=>s+asNum(m.amount),0);
  async function save(e) {
    e.preventDefault();
    if (asNum(form.amount) <= 0) return alert('El monto debe ser mayor a cero.');
    const { error } = await supabase.from('cash_movements').insert({ type: form.type, payment_method: form.method, amount: asNum(form.amount), note: form.note, store_id: profile?.store_id || DEFAULT_STORE_ID, user_id: profile?.id || null });
    if (error) alert(error.message); else { setForm({ type:'Ingreso', method:'Efectivo', amount:'0', note:'' }); reload(); }
  }
  const groupByType = todayMovs.reduce((acc, m) => { acc[m.type] = (acc[m.type] || 0) + asNum(m.amount); return acc; }, {});
  return (
    <div className="page">
      <div className="hero compact-hero"><h1>💰 Caja diaria</h1><p>Ingresos, egresos, compras, créditos y abonos del día.</p></div>
      <div className="kpi-grid">
        <Kpi label="Ingresos hoy" value={money(ingresos)} helper="Ventas, abonos y entradas" />
        <Kpi label="Egresos hoy" value={money(egresos)} helper="Compras, retiros y salidas" />
        <Kpi label="Créditos hoy" value={money(creditos)} helper="Ventas por cobrar" />
        <Kpi label="Caja neta" value={money(ingresos-egresos)} helper="Ingresos - egresos" />
      </div>
      <div className="two-col">
        <form className="card form-grid" onSubmit={save}>
          <label>Tipo<select value={form.type} onChange={e=>setForm({...form,type:e.target.value})}><option>Ingreso</option><option>Egreso</option><option>Apertura</option><option>Retiro</option><option>Compra</option></select></label>
          <label>Método<select value={form.method} onChange={e=>setForm({...form,method:e.target.value})}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option></select></label>
          <label>Monto<input value={form.amount} onChange={e=>setForm({...form,amount:e.target.value})} inputMode="decimal" /></label>
          <label>Nota<input value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Concepto del movimiento" /></label>
          <button className="primary-btn">Registrar movimiento</button>
        </form>
        <section className="card compact-card">
          <h3>Últimos movimientos</h3>
          {movements.slice(0,12).map(m=><div className="list-row" key={m.id}><span>{m.type} · {m.payment_method}<small>{fmtDate(m.created_at)} · {m.note || 'Sin nota'}</small></span><strong>{money(m.amount)}</strong></div>)}
          {!movements.length && <p className="muted">Aún no hay movimientos de caja.</p>}
        </section>
      </div>
      <div className="two-col extra-row">
        <section className="card compact-card"><h3>Resumen por tipo</h3>{Object.entries(groupByType).map(([k,v])=><div className="list-row" key={k}><span>{k}</span><strong>{money(v)}</strong></div>)}{!todayMovs.length && <p className="muted">Sin movimientos hoy.</p>}</section>
        <section className="card compact-card"><h3>Egresos y compras</h3>{egresosList.slice(0,10).map(m=><div className="list-row" key={m.id}><span>{m.type}<small>{fmtDate(m.created_at)} · {m.note || 'Sin nota'}</small></span><strong>{money(m.amount)}</strong></div>)}{!egresosList.length && <p className="muted">No hay egresos registrados hoy.</p>}</section>
      </div>
    </div>
  );
}


function Credits({ profile }) {
  const [sales, setSales] = useState([]);
  const [payments, setPayments] = useState([]);
  const [loading, setLoading] = useState(false);
  const [form, setForm] = useState({ sale_id:'', amount:'0', method:'Efectivo', note:'' });
  async function loadCredits() {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data: creditSales } = await supabase
      .from('sales')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .or('status.eq.Crédito,payment_method.eq.Crédito')
      .order('created_at', { ascending: false });
    setSales(creditSales || []);
    const { data: pays, error: payError } = await supabase
      .from('credit_payments')
      .select('*')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: false });
    if (!payError) setPayments(pays || []);
    setLoading(false);
  }
  useEffect(() => { loadCredits(); }, []);
  const paidBySale = payments.reduce((acc, p) => { acc[p.sale_id] = (acc[p.sale_id] || 0) + asNum(p.amount); return acc; }, {});
  const rows = sales.map(s => ({ ...s, paid: paidBySale[s.id] || 0, balance: Math.max(0, asNum(s.total) - (paidBySale[s.id] || 0)) })).filter(s => s.balance > 0.009);
  const totalPending = rows.reduce((s,v)=>s+asNum(v.balance),0);
  const selected = rows.find(r => r.id === form.sale_id) || rows[0];
  async function savePayment(e) {
    e.preventDefault();
    if (!selected) return alert('No hay crédito seleccionado.');
    const amount = Math.min(asNum(form.amount), asNum(selected.balance));
    if (amount <= 0) return alert('El abono debe ser mayor a cero.');
    const payload = { sale_id: selected.id, customer_name: selected.customer_name || 'Cliente', amount, method: form.method, note: form.note, store_id: profile?.store_id || DEFAULT_STORE_ID, user_id: profile?.id || null };
    const { error } = await supabase.from('credit_payments').insert(payload);
    if (error) return alert('Falta ejecutar la actualización SQL de V01.4: tabla credit_payments.');
    await supabase.from('cash_movements').insert({ type: 'Ingreso', payment_method: form.method, amount, note: `Abono crédito B${selected.receipt_number}. ${form.note || ''}`, store_id: profile?.store_id || DEFAULT_STORE_ID, user_id: profile?.id || null });
    if (amount >= asNum(selected.balance)) await supabase.from('sales').update({ status: 'Pagado' }).eq('id', selected.id);
    setForm({ sale_id:'', amount:'0', method:'Efectivo', note:'' });
    loadCredits();
  }
  const byClient = rows.reduce((acc, s) => { acc[s.customer_name || 'Cliente'] = (acc[s.customer_name || 'Cliente'] || 0) + asNum(s.balance); return acc; }, {});
  return (
    <div className="page">
      <div className="hero compact-hero"><h1>💳 Créditos</h1><p>Control de deuda, abonos y saldo real por cliente.</p></div>
      <div className="kpi-grid"><Kpi label="Saldo pendiente" value={money(totalPending)} helper={`${rows.length} comprobantes`} /><Kpi label="Clientes con deuda" value={Object.keys(byClient).length} helper="Cuentas activas" /><Kpi label="Abonado" value={money(payments.reduce((s,p)=>s+asNum(p.amount),0))} helper="Historial de abonos" /><Kpi label="Estado" value={loading ? 'Cargando' : 'Activo'} helper="Supabase" /></div>
      <div className="two-col">
        <section className="card compact-card"><h3>Deuda por cliente</h3>{Object.entries(byClient).map(([client, amount])=><div className="list-row" key={client}><span>{client}</span><strong>{money(amount)}</strong></div>)}{!rows.length && <p className="muted">No hay créditos pendientes.</p>}</section>
        <form className="card form-grid" onSubmit={savePayment}>
          <h3>Registrar abono</h3>
          <label>Crédito<select value={form.sale_id || selected?.id || ''} onChange={e=>setForm({...form,sale_id:e.target.value})}>{rows.map(s=><option key={s.id} value={s.id}>B{s.receipt_number} · {s.customer_name} · Saldo {money(s.balance)}</option>)}</select></label>
          <label>Monto del abono<input value={form.amount} onChange={e=>setForm({...form,amount:e.target.value})} inputMode="decimal" placeholder="Monto recibido" /></label>
          <label>Método<select value={form.method} onChange={e=>setForm({...form,method:e.target.value})}><option>Efectivo</option><option>Yape</option><option>Plin</option><option>Transferencia</option><option>Tarjeta</option></select></label>
          <label>Nota<input value={form.note} onChange={e=>setForm({...form,note:e.target.value})} placeholder="Observación del abono" /></label>
          <button className="primary-btn" disabled={!rows.length}>Guardar abono</button>
        </form>
      </div>
      <section className="card compact-card extra-row"><h3>Comprobantes pendientes</h3>{rows.map(s=><div className="list-row" key={s.id}><span>B{s.receipt_number} · {s.customer_name || 'Cliente'}<small>Total {money(s.total)} · Abonado {money(s.paid)} · {fmtDate(s.created_at)}</small></span><strong>{money(s.balance)}</strong></div>)}{!rows.length && <p className="muted">Sin comprobantes pendientes.</p>}</section>
    </div>
  );
}


function Reports({ products, profile }) {
  const { sales } = useSales(profile);
  const { movements } = useCashMovements(profile);
  const [payments, setPayments] = useState([]);
  useEffect(() => { if (hasSupabaseConfig) supabase.from('credit_payments').select('*').eq('store_id', profile?.store_id || DEFAULT_STORE_ID).then(({data}) => setPayments(data || [])); }, [profile?.store_id]);
  const totalSales = sales.reduce((s,v)=>s+asNum(v.total),0);
  const creditSales = sales.filter(s => s.status === 'Crédito' || s.payment_method === 'Crédito');
  const totalCredits = creditSales.reduce((s,v)=>s+asNum(v.total),0);
  const totalPaidCredits = payments.reduce((s,p)=>s+asNum(p.amount),0);
  const avg = sales.length ? totalSales / sales.length : 0;
  const byMethod = sales.reduce((acc, s) => { acc[s.payment_method || 'Sin método'] = (acc[s.payment_method || 'Sin método'] || 0) + asNum(s.total); return acc; }, {});
  const byCat = products.reduce((acc, p) => { acc[p.category || 'General'] = (acc[p.category || 'General'] || 0) + asNum(p.stock); return acc; }, {});
  const egresos = movements.filter(m => ['Egreso','Compra','Retiro','Compra crédito'].includes(m.type)).reduce((s,m)=>s+asNum(m.amount),0);
  return <div className="page"><div className="hero compact-hero"><h1>📈 Reportes</h1><p>Ventas, caja, créditos, abonos e inventario en vista rápida.</p></div><div className="kpi-grid"><Kpi label="Total vendido" value={money(totalSales)} helper={`${sales.length} ventas`} /><Kpi label="Ticket promedio" value={money(avg)} helper="Promedio de venta" /><Kpi label="Crédito pendiente" value={money(Math.max(0,totalCredits-totalPaidCredits))} helper="Por cobrar" /><Kpi label="Egresos" value={money(egresos)} helper="Compras y salidas" /></div><div className="two-col"><section className="card compact-card"><h3>Ventas por método</h3>{Object.entries(byMethod).map(([k,v])=><div className="list-row" key={k}><span>{k}</span><strong>{money(v)}</strong></div>)}{!sales.length && <p className="muted">No hay ventas todavía.</p>}</section><section className="card compact-card"><h3>Stock por categoría</h3>{Object.entries(byCat).map(([k,v])=><div className="list-row" key={k}><span>{k}</span><strong>{v}</strong></div>)}</section></div><div className="two-col extra-row"><section className="card compact-card"><h3>Abonos recibidos</h3>{payments.slice(0,8).map(p=><div className="list-row" key={p.id}><span>{p.customer_name}<small>{fmtDate(p.created_at)} · {p.method}</small></span><strong>{money(p.amount)}</strong></div>)}{!payments.length && <p className="muted">No hay abonos registrados.</p>}</section><section className="card compact-card"><h3>Movimientos de caja</h3>{movements.slice(0,8).map(m=><div className="list-row" key={m.id}><span>{m.type} · {m.payment_method}<small>{m.note || 'Sin nota'}</small></span><strong>{money(m.amount)}</strong></div>)}</section></div></div>;
}



function UsersAdmin({ profile }) {
  const [profiles, setProfiles] = useState([]);
  const [saving, setSaving] = useState('');
  const [drafts, setDrafts] = useState({});
  const roles = ['dueno', 'admin', 'cajero', 'almacen', 'lectura'];

  async function loadProfiles() {
    const { data, error } = await supabase
      .from('profiles')
      .select('id,email,full_name,role,status,created_at,store_id')
      .eq('store_id', profile?.store_id || DEFAULT_STORE_ID)
      .order('created_at', { ascending: true });
    if (!error) setProfiles(data || []);
  }

  useEffect(() => { loadProfiles(); }, [profile?.store_id]);

  function draft(row) {
    return drafts[row.id] || { full_name: row.full_name || '', role: row.role || 'cajero', status: row.status || 'Activo' };
  }

  function setDraftValue(row, key, value) {
    setDrafts(prev => ({ ...prev, [row.id]: { ...draft(row), [key]: value } }));
  }

  async function saveProfile(row) {
    const next = draft(row);
    setSaving(row.id);
    const { error } = await supabase
      .from('profiles')
      .update({ full_name: next.full_name, role: next.role, status: next.status, updated_at: new Date().toISOString() })
      .eq('id', row.id);
    setSaving('');
    if (error) return alert(error.message || 'No se pudo actualizar el usuario. Revisa políticas SQL V01.6.');
    await loadProfiles();
  }

  return (
    <div className="page">
      <div className="hero compact-hero"><h1>🧑‍💼 Usuarios y roles</h1><p>Administra permisos por usuario. Los usuarios nuevos se crean primero en Supabase Auth.</p></div>
      <section className="card compact-card">
        <h3>Cómo crear un usuario nuevo</h3>
        <p className="muted">Ruta: Supabase → Authentication → Users → Add user → Create new user. Activa Auto Confirm User. Luego regresa aquí para asignar rol.</p>
      </section>
      <section className="card compact-card extra-row">
        <h3>Usuarios registrados</h3>
        {profiles.map(row => {
          const d = draft(row);
          return (
            <div className="list-row" key={row.id}>
              <span>
                <strong>{row.email || 'Sin correo'}</strong>
                <small>Creado: {fmtDate(row.created_at)} · Rol actual: {ROLE_LABELS[row.role] || row.role}</small>
              </span>
              <input value={d.full_name} onChange={e=>setDraftValue(row, 'full_name', e.target.value)} placeholder="Nombre" />
              <select value={d.role} onChange={e=>setDraftValue(row, 'role', e.target.value)}>
                {roles.map(r => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
              </select>
              <select value={d.status} onChange={e=>setDraftValue(row, 'status', e.target.value)}>
                <option>Activo</option><option>Inactivo</option>
              </select>
              <button className="secondary-btn" onClick={() => saveProfile(row)} disabled={saving === row.id}>{saving === row.id ? 'Guardando...' : 'Guardar'}</button>
            </div>
          );
        })}
        {!profiles.length && <p className="muted">No hay perfiles registrados todavía.</p>}
      </section>
    </div>
  );
}


function LabelsAdmin({ products = [], categories = [], subcategories = [], store }) {
  const [search, setSearch] = useState('');
  const [categoryId, setCategoryId] = useState('all');
  const [subcategoryId, setSubcategoryId] = useState('all');
  const [mode, setMode] = useState('both');
  const [labelSize, setLabelSize] = useState('medium');
  const [showPrice, setShowPrice] = useState(true);
  const [showLogo, setShowLogo] = useState(true);
  const [selected, setSelected] = useState(new Set());

  const activeProducts = useMemo(() => products.filter(p => p.active !== false && p.status !== 'Inactivo'), [products]);
  const filtered = useMemo(() => activeProducts.filter(p => {
    const q = normalizeText(search);
    const matchText = !q || [p.name, p.code, p.barcode, p.brand, p.color, p.category, p.subcategory].some(v => normalizeText(v).includes(q));
    const matchCat = categoryId === 'all' || p.category_id === categoryId || normalizeText(p.category) === normalizeText(categories.find(c=>c.id===categoryId)?.name);
    const matchSub = subcategoryId === 'all' || p.subcategory_id === subcategoryId || normalizeText(p.subcategory) === normalizeText(subcategories.find(c=>c.id===subcategoryId)?.name);
    return matchText && matchCat && matchSub;
  }), [activeProducts, search, categoryId, subcategoryId, categories, subcategories]);

  const chosen = useMemo(() => activeProducts.filter(p => selected.has(p.id)), [activeProducts, selected]);
  const printable = chosen.length ? chosen : filtered;
  const visibleSubcategories = useMemo(() => categoryId === 'all' ? subcategories : subcategories.filter(s => s.parent_id === categoryId), [categoryId, subcategories]);

  function toggleProduct(id) {
    setSelected(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }
  function selectFiltered() { setSelected(new Set(filtered.map(p => p.id))); }
  function clearSelection() { setSelected(new Set()); }
  function printLabels() {
    if (!printable.length) return alert('No hay productos para imprimir.');
    setTimeout(() => window.print(), 100);
  }

  return (
    <div className="page labels-page">
      <div className="hero compact-hero"><h1>🏷️ Etiquetas QR y código de barras</h1><p>Genera etiquetas en PDF para imprimir, recortar y pegar en productos.</p></div>
      <div className="tool-grid labels-tools">
        <section className="card compact-card">
          <h3>Filtros</h3>
          <label>Buscar producto<input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Nombre, código, barcode, marca o color" /></label>
          <label>Categoría<select value={categoryId} onChange={e=>{ setCategoryId(e.target.value); setSubcategoryId('all'); }}><option value="all">Todas</option>{categories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></label>
          <label>Subcategoría<select value={subcategoryId} onChange={e=>setSubcategoryId(e.target.value)}><option value="all">Todas</option>{visibleSubcategories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></label>
          <div className="import-summary label-summary"><Kpi label="Filtrados" value={filtered.length} helper="productos" /><Kpi label="Seleccionados" value={chosen.length || filtered.length} helper={chosen.length ? 'manual' : 'por filtro'} /></div>
          <div className="button-row"><button className="secondary-btn" onClick={selectFiltered}>Seleccionar filtrados</button><button className="secondary-btn" onClick={clearSelection}>Limpiar selección</button></div>
        </section>

        <section className="card compact-card">
          <h3>Diseño de etiqueta</h3>
          <label>Tipo de código<select value={mode} onChange={e=>setMode(e.target.value)}><option value="both">QR + barras</option><option value="qr">Solo QR</option><option value="barcode">Solo código de barras</option></select></label>
          <label>Tamaño<select value={labelSize} onChange={e=>setLabelSize(e.target.value)}><option value="small">Pequeña 40 x 30 mm</option><option value="medium">Mediana 50 x 30 mm</option><option value="large">Ropa 60 x 40 mm</option></select></label>
          <label className="check-row"><input type="checkbox" checked={showPrice} onChange={e=>setShowPrice(e.target.checked)} /> Mostrar precio</label>
          <label className="check-row"><input type="checkbox" checked={showLogo} onChange={e=>setShowLogo(e.target.checked)} /> Mostrar marca Clomar Store</label>
          <button className="primary-btn" onClick={printLabels}>Imprimir / Guardar PDF</button>
          <p className="muted">En la ventana de impresión elige <strong>Guardar como PDF</strong> o tu impresora de etiquetas.</p>
        </section>
      </div>

      <section className="card compact-card">
        <h3>Productos para etiquetas</h3>
        <div className="product-pick-list">
          {filtered.map(p => <label key={p.id} className="product-pick-row"><input type="checkbox" checked={selected.has(p.id)} onChange={()=>toggleProduct(p.id)} /><img src={productImageSrc(p)} alt={p.name}/><span><strong>{p.name}</strong><small>{p.code} · {p.barcode || 'Sin barcode'} · {p.category || 'Sin categoría'}{p.subcategory ? ` / ${p.subcategory}` : ''}</small></span><b>{money(p.price)}</b></label>)}
          {!filtered.length && <p className="muted">No hay productos con esos filtros.</p>}
        </div>
      </section>

      <section className="card compact-card no-print">
        <h3>Vista previa</h3>
        <p className="muted">Se imprimen {printable.length} etiquetas. Si no seleccionas productos, se imprimen todos los filtrados.</p>
      </section>

      <div className={`print-label-sheet label-size-${labelSize}`}>
        {printable.map(product => {
          const code = productScanCode(product);
          return <div className="print-label" key={product.id}>
            {showLogo && <div className="label-brand"><img src={APP_ICON} alt="Clomar"/><span>{store?.name || 'Clomar Store'}</span></div>}
            <div className="label-name">{product.name}</div>
            {showPrice && <div className="label-price">{money(product.price)}</div>}
            <div className={`label-codes mode-${mode}`}>{(mode === 'qr' || mode === 'both') && <img className="label-qr" src={qrUrl(code)} alt={`QR ${code}`} />}{(mode === 'barcode' || mode === 'both') && <BarcodeSVG value={code} />}</div>
            <div className="label-code-text">{code}</div>
          </div>;
        })}
      </div>
    </div>
  );
}

function ToolsAdmin({ profile, products = [], categories = [], subcategories = [], reloadProducts, reloadCustomers }) {
  const [confirmText, setConfirmText] = useState('');
  const [resetting, setResetting] = useState(false);
  const [deleteImages, setDeleteImages] = useState(false);
  const [fileName, setFileName] = useState('');
  const [previewRows, setPreviewRows] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState(null);

  const categoryByName = useMemo(() => {
    const map = new Map();
    categories.forEach(c => map.set(normalizeText(c.name), c));
    return map;
  }, [categories]);
  const subcategoryByParentAndName = useMemo(() => {
    const map = new Map();
    subcategories.forEach(sc => map.set(`${sc.parent_id}|${normalizeText(sc.name)}`, sc));
    return map;
  }, [subcategories]);

  function normalizeImportedRow(row, idx) {
    const normalized = {};
    Object.entries(row || {}).forEach(([key, value]) => {
      const field = canonicalProductField(normalizeHeader(key));
      normalized[field] = typeof value === 'string' ? value.trim() : value;
    });
    const categoryName = String(normalized.category || '').trim();
    const subcategoryName = String(normalized.subcategory || '').trim();
    const category = categoryByName.get(normalizeText(categoryName));
    const subcategory = category ? subcategoryByParentAndName.get(`${category.id}|${normalizeText(subcategoryName)}`) : null;
    const generatedCode = `${categoryPrefix(categoryName)}-${String(products.length + idx + 1).padStart(6, '0')}`;
    const code = String(normalized.code || '').trim() || generatedCode;
    const barcode = String(normalized.barcode || '').trim() || code;
    const errors = [];
    if (!String(normalized.name || '').trim()) errors.push('Falta nombre');
    if (!categoryName) errors.push('Falta categoría');
    if (categoryName && !category) errors.push(`Categoría no existe: ${categoryName}`);
    if (subcategoryName && category && !subcategory) errors.push(`Subcategoría no existe: ${subcategoryName}`);
    if (parseMoneyLike(normalized.price) <= 0) errors.push('Precio debe ser mayor a 0');
    return {
      rowNumber: idx + 2,
      code,
      barcode,
      name: String(normalized.name || '').trim(),
      category: category?.name || categoryName || 'General',
      subcategory: subcategory?.name || subcategoryName || '',
      category_id: category?.id || null,
      subcategory_id: subcategory?.id || null,
      brand: String(normalized.brand || '').trim(),
      size: String(normalized.size || '').trim(),
      color: String(normalized.color || '').trim(),
      description: String(normalized.description || '').trim(),
      cost: parseMoneyLike(normalized.cost),
      price: parseMoneyLike(normalized.price),
      stock: parseMoneyLike(normalized.stock),
      stock_min: parseMoneyLike(normalized.stock_min || 1),
      image_url: String(normalized.image_url || '').trim(),
      active: boolActive(normalized.active ?? true),
      errors,
    };
  }

  async function parseFile(e) {
    const file = e.target.files?.[0];
    setImportResult(null);
    setPreviewRows([]);
    setRawRows([]);
    if (!file) return;
    setFileName(file.name);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const normalized = rows.map((r, idx) => normalizeImportedRow(r, idx));
      setRawRows(normalized);
      setPreviewRows(normalized.slice(0, 30));
    } catch (error) {
      alert(error.message || 'No se pudo leer el archivo Excel.');
    }
  }

  async function importProducts() {
    if (!rawRows.length) return alert('Primero carga un Excel.');
    const invalid = rawRows.filter(r => r.errors.length);
    if (invalid.length) return alert(`Hay ${invalid.length} filas con errores. Corrige el Excel antes de importar.`);
    setImporting(true);
    try {
      const payloads = rawRows.map(r => ({
        code: r.code,
        barcode: r.barcode,
        name: r.name,
        category: r.category,
        subcategory: r.subcategory,
        category_id: r.category_id,
        subcategory_id: r.subcategory_id,
        brand: r.brand,
        size: r.size,
        color: r.color,
        description: r.description,
        cost: r.cost,
        price: r.price,
        stock: r.stock,
        stock_min: r.stock_min,
        image_url: r.image_url,
        image_path: '',
        active: r.active,
        status: r.active ? 'Activo' : 'Inactivo',
        store_id: profile?.store_id || DEFAULT_STORE_ID,
        created_by: profile?.id || null,
        updated_at: new Date().toISOString(),
      }));
      const chunks = [];
      for (let i = 0; i < payloads.length; i += 200) chunks.push(payloads.slice(i, i + 200));
      let count = 0;
      for (const chunk of chunks) {
        const { error } = await supabase.from('products').upsert(chunk, { onConflict: 'code' });
        if (error) throw error;
        count += chunk.length;
      }
      await supabase.from('product_import_batches').insert({
        store_id: profile?.store_id || DEFAULT_STORE_ID,
        user_id: profile?.id || null,
        file_name: fileName,
        total_rows: rawRows.length,
        imported_rows: count,
        status: 'Importado',
        notes: 'Importación desde V01.10',
      }).then(()=>{});
      setImportResult({ count });
      setPreviewRows([]);
      setRawRows([]);
      setFileName('');
      await reloadProducts?.();
      alert(`Importación completada: ${count} productos.`);
    } catch (error) {
      alert(error.message || 'No se pudo importar productos.');
    } finally {
      setImporting(false);
    }
  }

  async function deleteProductImagesFromStorage() {
    try {
      const { data } = await supabase.storage.from('product-images').list('', { limit: 1000 });
      const files = (data || []).map(f => f.name).filter(Boolean);
      if (files.length) await supabase.storage.from('product-images').remove(files);
    } catch (_) {}
  }

  async function controlledReset() {
    if (profile?.role !== 'dueno') return alert('Solo el dueño puede reiniciar datos.');
    if (confirmText !== 'REINICIAR CLOMAR') return alert('Debes escribir exactamente: REINICIAR CLOMAR');
    if (!confirm('Esto borrará productos, clientes, ventas, caja, créditos y movimientos de prueba. Se conservan usuarios, tienda y categorías. ¿Continuar?')) return;
    setResetting(true);
    try {
      const { error } = await supabase.rpc('clomar_reset_operational_data', { confirm_text: confirmText, store_uuid: profile?.store_id || DEFAULT_STORE_ID });
      if (error) throw error;
      if (deleteImages) await deleteProductImagesFromStorage();
      setConfirmText('');
      await reloadProducts?.();
      await reloadCustomers?.();
      alert('Reinicio controlado completado. El sistema quedó listo para cargar productos reales.');
    } catch (error) {
      alert(error.message || 'No se pudo reiniciar. Verifica que ejecutaste el SQL V01.10.');
    } finally {
      setResetting(false);
    }
  }

  return (
    <div className="page">
      <div className="hero compact-hero"><h1>🛠️ Herramientas</h1><p>Reinicio controlado, importación desde Excel y preparación para productos reales.</p></div>
      <div className="tool-grid">
        <section className="card compact-card danger-zone-card">
          <h3>Reinicio controlado</h3>
          <p className="muted">Borra datos operativos de prueba y conserva usuarios, roles, tienda, logo, categorías y subcategorías.</p>
          <div className="reset-keep-delete">
            <div><strong>Conserva</strong><small>Usuarios · Roles · Tienda · Logo · Categorías</small></div>
            <div><strong>Borra</strong><small>Productos · Ventas · Caja · Créditos · Clientes · Movimientos</small></div>
          </div>
          <label className="check-row"><input type="checkbox" checked={deleteImages} onChange={e=>setDeleteImages(e.target.checked)} /> Borrar también imágenes del bucket product-images</label>
          <label>Confirmación obligatoria<input value={confirmText} onChange={e=>setConfirmText(e.target.value)} placeholder="Escribe REINICIAR CLOMAR" /></label>
          <button className="danger-btn" disabled={resetting} onClick={controlledReset}>{resetting ? 'Reiniciando...' : 'Reiniciar datos de prueba'}</button>
        </section>

        <section className="card compact-card">
          <h3>Importar productos desde Excel</h3>
          <p className="muted">Carga una plantilla .xlsx con productos reales. Si falta código o barcode, la app genera uno interno.</p>
          <a className="secondary-btn" href="/plantilla_productos_clomar_v0110.xlsx" download>Descargar plantilla Excel</a>
          <label>Seleccionar archivo Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={parseFile} /></label>
          {fileName && <div className="info-box">Archivo cargado: <strong>{fileName}</strong> · Filas leídas: {rawRows.length}</div>}
          {rawRows.length > 0 && <div className="import-summary">
            <Kpi label="Filas" value={rawRows.length} helper="productos detectados" />
            <Kpi label="Errores" value={rawRows.filter(r=>r.errors.length).length} helper="deben corregirse" />
            <Kpi label="Listos" value={rawRows.filter(r=>!r.errors.length).length} helper="para importar" />
          </div>}
          {previewRows.length > 0 && <div className="preview-table-wrap">
            <table className="preview-table"><thead><tr><th>Fila</th><th>Código</th><th>Producto</th><th>Categoría</th><th>Subcategoría</th><th>Precio</th><th>Stock</th><th>Estado</th></tr></thead><tbody>
              {previewRows.map(r => <tr key={r.rowNumber} className={r.errors.length ? 'row-error' : ''}><td>{r.rowNumber}</td><td>{r.code}</td><td>{r.name}</td><td>{r.category}</td><td>{r.subcategory}</td><td>{money(r.price)}</td><td>{r.stock}</td><td>{r.errors.length ? r.errors.join('; ') : 'OK'}</td></tr>)}
            </tbody></table>
          </div>}
          <button className="primary-btn" disabled={!rawRows.length || importing || rawRows.some(r=>r.errors.length)} onClick={importProducts}>{importing ? 'Importando...' : 'Importar productos'}</button>
          {importResult && <div className="success-box">Importación completada: {importResult.count} productos.</div>}
        </section>
      </div>
    </div>
  );
}

function StoreSettings({ store, reloadProfile }) {
  const [form, setForm] = useState(store || {});
  const [saving, setSaving] = useState(false);

  useEffect(() => { setForm(store || {}); }, [store?.id]);

  async function save(e) {
    e.preventDefault();
    setSaving(true);
    const payload = {
      name: form.name || 'Clomar Store Pro',
      ruc: form.ruc || '',
      address: form.address || '',
      phone: form.phone || '',
      email: form.email || '',
      logo_url: form.logo_url || '',
      updated_at: new Date().toISOString(),
    };
    const { error } = await supabase.from('stores').update(payload).eq('id', store?.id || DEFAULT_STORE_ID);
    setSaving(false);
    if (error) return alert(error.message || 'No se pudo actualizar la tienda.');
    await reloadProfile();
    alert('Datos de tienda actualizados.');
  }

  return (
    <div className="page">
      <div className="hero compact-hero"><h1>🏪 Configuración de tienda</h1><p>Datos base para usuarios, tickets, comprobantes y catálogo.</p></div>
      <form className="card form-grid" onSubmit={save}>
        <label>Nombre comercial<input value={form.name || ''} onChange={e=>setForm({...form, name:e.target.value})} placeholder="Clomar Store Pro" /></label>
        <label>RUC<input value={form.ruc || ''} onChange={e=>setForm({...form, ruc:e.target.value})} placeholder="RUC de la tienda" /></label>
        <label>Dirección<input value={form.address || ''} onChange={e=>setForm({...form, address:e.target.value})} placeholder="Dirección" /></label>
        <label>Teléfono / WhatsApp<input value={form.phone || ''} onChange={e=>setForm({...form, phone:e.target.value})} placeholder="Celular" /></label>
        <label>Correo<input value={form.email || ''} onChange={e=>setForm({...form, email:e.target.value})} placeholder="correo@tienda.com" /></label>
        <label>Logo URL<input value={form.logo_url || ''} onChange={e=>setForm({...form, logo_url:e.target.value})} placeholder="Opcional: pega una URL pública del logo" /></label>
        <div className="brand-preview">
          <span>Vista de marca</span>
          <img src={form.logo_url || APP_LOGO_FULL} alt="Logo Clomar Store" />
        </div>
        <button className="primary-btn" disabled={saving}>{saving ? 'Guardando...' : 'Guardar tienda'}</button>
      </form>
    </div>
  );
}

function AppShell({ session }) {
  const [current, setCurrent] = useState('ventas');
  const [open, setOpen] = useState(false);
  const { profile, store, loading: profileLoading, reload: reloadProfile } = useUserProfile(session);
  const { products, loading, reload } = useProducts(profile);
  const { customers, reload: reloadCustomers } = useCustomers(profile);
  const { categories, subcategories, reload: reloadCategories } = useCategories(profile);

  useEffect(() => {
    if (profile && !canAccess(profile.role, current)) setCurrent(firstAllowedModule(profile.role));
  }, [profile?.role, current]);

  if (profileLoading) return <div className="loader full">Cargando perfil y permisos...</div>;
  if (profile?.status === 'Inactivo') return <InactiveUser profile={profile} />;

  const contentMap = {
    panel: <Panel products={products} profile={profile}/>,
    ventas: <POS products={products} reloadProducts={reload} customers={customers} profile={profile}/>,
    productos: <Products products={products} reload={reload} profile={profile} categories={categories} subcategories={subcategories} reloadCategories={reloadCategories}/>,
    categorias: <CategoriesAdmin profile={profile} categories={categories} subcategories={subcategories} products={products} reloadCategories={reloadCategories}/>,
    etiquetas: <LabelsAdmin products={products} categories={categories} subcategories={subcategories} store={store}/>,
    inventario: <Inventory products={products}/>,
    reportes: <Reports products={products} profile={profile}/>,
    creditos: <Credits profile={profile}/>,
    caja: <CashPage profile={profile} />,
    ingreso: <StockEntry products={products} reloadProducts={reload} profile={profile}/>,
    clientes: <Customers customers={customers} reload={reloadCustomers} profile={profile}/>,
    usuarios: <UsersAdmin profile={profile}/>,
    tienda: <StoreSettings store={store} reloadProfile={reloadProfile}/>,
    herramientas: <ToolsAdmin profile={profile} products={products} categories={categories} subcategories={subcategories} reloadProducts={reload} reloadCustomers={reloadCustomers}/>,
  };
  const content = canAccess(profile?.role, current) ? contentMap[current] : <AccessDenied profile={profile} setCurrent={setCurrent} />;
  return <div className="app"><Sidebar current={current} setCurrent={setCurrent} open={open} setOpen={setOpen} session={session} profile={profile} store={store}/><main className="main"><Header setOpen={setOpen} current={current} profile={profile} store={store}/>{loading ? <div className="loader">Cargando...</div> : content}</main></div>;
}

function Root() {
  const { session, loading } = useAuth();
  useEffect(() => {
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(()=>{});
  }, []);
  if (loading) return <div className="loader full">Iniciando Clomar Store...</div>;
  if (!hasSupabaseConfig) return <Login />;
  return session ? <AppShell session={session}/> : <Login/>;
}

createRoot(document.getElementById('root')).render(<Root />);
