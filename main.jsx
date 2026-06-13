import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { supabase, hasSupabaseConfig } from './supabaseClient';
import { LogOut, Menu, Search, ShoppingCart, X } from 'lucide-react';
import './styles.css';

const money = (value) => `S/ ${Number(value || 0).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const todayISO = () => new Date().toISOString().slice(0, 10);
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
  { id: 'demo-1', code: '0001', name: 'Zapatillas deportivas Newton Nimble Leather', category: 'Calzado', price: 380, cost: 220, stock: 5, stock_min: 2, image_url: '' },
  { id: 'demo-2', code: '0002', name: 'Sombrero para el sol Bora Bora Booney', category: 'Accesorios', price: 80, cost: 40, stock: 2, stock_min: 2, image_url: '' },
  { id: 'demo-3', code: '0003', name: 'Camisa de popelina de manga larga para hombre', category: 'Ropa', price: 120, cost: 55, stock: 1, stock_min: 2, image_url: '' },
];

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
        <div className="brand-logo">🛍️</div>
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

function Sidebar({ current, setCurrent, open, setOpen, session }) {
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
      ['inventario', '📘', 'Inventario'],
      ['ingreso', '📥', 'Ingreso mercadería'],
    ]},
    { title: 'Contactos', items: [['clientes', '👥', 'Clientes']]},
  ];
  return (
    <aside className={`sidebar ${open ? 'open' : ''}`}>
      <div className="sidebar-head">
        <div className="mini-logo">🛍️</div>
        <div>
          <strong>Clomar Store Pro</strong>
          <small>{session?.user?.email || 'Usuario'}</small>
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

function Header({ setOpen, current }) {
  const titleMap = {
    panel: 'Panel dueño', ventas: 'Venta rápida', creditos: 'Créditos', caja: 'Caja diaria', reportes: 'Reportes', productos: 'Productos', inventario: 'Inventario', ingreso: 'Ingreso de mercadería', clientes: 'Clientes'
  };
  return (
    <header className="app-header">
      <button className="ghost mobile-only" onClick={() => setOpen(true)}><Menu/></button>
      <div>
        <h2>{titleMap[current]}</h2>
        <p>Clomar Store Pro · App comercial rápida</p>
      </div>
    </header>
  );
}

function useProducts() {
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
      .select('id,code,name,category,price,cost,stock,image_url,status,stock_min')
      .eq('status', 'Activo')
      .order('name');
    if (!error) setProducts(data || []);
    setLoading(false);
  }

  useEffect(() => { loadProducts(); }, []);
  return { products, loading, reload: loadProducts };
}

function useCustomers() {
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
      .order('name');
    if (!error) setCustomers(data || []);
    setLoading(false);
  }

  useEffect(() => { loadCustomers(); }, []);
  return { customers, loading, reload: loadCustomers };
}

function useSales() {
  const [sales, setSales] = useState([]);
  const [loading, setLoading] = useState(false);
  async function loadSales(limit = 50) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data, error } = await supabase
      .from('sales')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(limit);
    if (!error) setSales(data || []);
    setLoading(false);
  }
  useEffect(() => { loadSales(); }, []);
  return { sales, loading, reload: loadSales };
}

function useCashMovements() {
  const [movements, setMovements] = useState([]);
  const [loading, setLoading] = useState(false);
  async function loadMovements(limit = 80) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data, error } = await supabase
      .from('cash_movements')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(limit);
    if (!error) setMovements(data || []);
    setLoading(false);
  }
  useEffect(() => { loadMovements(); }, []);
  return { movements, loading, reload: loadMovements };
}

function useStockMovements() {
  const [movements, setMovements] = useState([]);
  const [loading, setLoading] = useState(false);
  async function loadMovements(limit = 80) {
    if (!hasSupabaseConfig) return;
    setLoading(true);
    const { data, error } = await supabase
      .from('stock_movements')
      .select('*, products(name, code)')
      .order('created_at', { ascending: false })
      .limit(limit);
    if (!error) setMovements(data || []);
    setLoading(false);
  }
  useEffect(() => { loadMovements(); }, []);
  return { movements, loading, reload: loadMovements };
}

function Kpi({ label, value, helper }) {
  return <div className="kpi"><span>{label}</span><strong>{value}</strong><small>{helper}</small></div>;
}

function Panel({ products }) {
  const { sales } = useSales();
  const { movements } = useCashMovements();
  const stockCritico = products.filter(p => asNum(p.stock) <= asNum(p.stock_min ?? 2));
  const today = todayISO();
  const salesToday = sales.filter(s => String(s.created_at || '').slice(0,10) === today);
  const ingresosToday = movements.filter(m => String(m.created_at || '').slice(0,10) === today && ['Ingreso','Apertura'].includes(m.type));
  const egresosToday = movements.filter(m => String(m.created_at || '').slice(0,10) === today && ['Egreso','Compra','Retiro'].includes(m.type));
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

function POS({ products, reloadProducts, customers }) {
  const [query, setQuery] = useState('');
  const [cart, setCart] = useState([]);
  const [method, setMethod] = useState('Efectivo');
  const [customer, setCustomer] = useState('Cliente');
  const [saving, setSaving] = useState(false);
  const [lastTicket, setLastTicket] = useState(null);
  const normalized = query.trim().toLowerCase();
  const matches = useMemo(() => {
    const base = !normalized ? products.slice(0, 12) : products.filter(p => `${p.code} ${p.name} ${p.category}`.toLowerCase().includes(normalized)).slice(0, 20);
    return base;
  }, [products, normalized]);
  const total = cart.reduce((sum, item) => sum + asNum(item.price) * asNum(item.qty), 0);

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
    const salePayload = { customer_name: customer || 'Cliente', payment_method: method, total, status: method === 'Crédito' ? 'Crédito' : 'Pagado' };
    const { data: sale, error } = await supabase.from('sales').insert(salePayload).select().single();
    if (error) { alert(error.message); setSaving(false); return; }
    const items = cart.map(item => ({ sale_id: sale.id, product_id: item.id, product_name: item.name, qty: item.qty, price: item.price, subtotal: asNum(item.qty) * asNum(item.price) }));
    await supabase.from('sale_items').insert(items);
    for (const item of cart) {
      await supabase.from('products').update({ stock: asNum(item.stock) - asNum(item.qty) }).eq('id', item.id);
      await supabase.from('stock_movements').insert({ product_id: item.id, type: 'Salida', qty: item.qty, note: `Venta B${sale.receipt_number || sale.id}` });
    }
    await supabase.from('cash_movements').insert({ type: method === 'Crédito' ? 'Crédito' : 'Ingreso', payment_method: method, amount: total, note: `Venta B${sale.receipt_number || sale.id}` });
    setLastTicket({ sale, items });
    setCart([]); setSaving(false); reloadProducts();
  }

  return (
    <div className="page pos-page">
      <div className="hero compact-hero"><h1>🧾 Venta rápida</h1><p>Buscador instantáneo, carrito fijo, cliente, crédito y comprobante.</p></div>
      <div className="pos-layout">
        <section className="card compact-card">
          <div className="search-box"><Search size={18}/><input value={query} onChange={(e)=>setQuery(e.target.value)} placeholder="Buscar por nombre, código o categoría..." autoFocus /></div>
          <div className="product-list">
            {matches.map(product => (
              <button key={product.id} className="product-row" onClick={() => addProduct(product)}>
                <div><strong>{product.name}</strong><small>{product.code} · {product.category} · Stock {asNum(product.stock)}</small></div>
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


function Products({ products, reload }) {
  const [form, setForm] = useState({ code:'', name:'', category:'General', price:'', cost:'', stock:'0', stock_min:'2' });
  const fieldLabels = {
    code: 'Código', name: 'Nombre del producto', category: 'Categoría', price: 'Precio de venta', cost: 'Costo de compra', stock: 'Stock inicial', stock_min: 'Stock mínimo'
  };
  const fieldPlaceholders = {
    code: 'Ejemplo: 0004 o código de barras', name: 'Nombre comercial del producto', category: 'Ejemplo: Ropa, Calzado, Accesorios', price: 'Precio al cliente', cost: 'Costo de compra', stock: 'Cantidad disponible', stock_min: 'Alerta mínima'
  };
  async function save(e) {
    e.preventDefault();
    if (!hasSupabaseConfig) return alert('Configura Supabase para guardar productos.');
    const payload = { ...form, price:asNum(form.price), cost:asNum(form.cost), stock:asNum(form.stock), stock_min:asNum(form.stock_min), status:'Activo' };
    const { error } = await supabase.from('products').insert(payload);
    if (error) alert(error.message); else { setForm({ code:'', name:'', category:'General', price:'', cost:'', stock:'0', stock_min:'2' }); reload(); }
  }
  return <div className="page"><div className="hero compact-hero"><h1>📦 Productos</h1><p>Crea artículos para vender y controlar stock.</p></div><div className="two-col"><form className="card form-grid" onSubmit={save}>{['code','name','category','price','cost','stock','stock_min'].map(k=><label key={k}>{fieldLabels[k]}<input value={form[k]} placeholder={fieldPlaceholders[k]} inputMode={['price','cost','stock','stock_min'].includes(k) ? 'decimal' : 'text'} onChange={e=>setForm({...form,[k]:e.target.value})}/></label>)}<button className="primary-btn">Guardar producto</button></form><section className="card compact-card"><h3>Lista de productos</h3>{products.map(p=><div className="list-row" key={p.id}><span>{p.code} · {p.name}</span><strong>{money(p.price)}</strong></div>)}</section></div></div>;
}

function Customers({ customers, reload }) {
  const [form, setForm] = useState({ name:'', phone:'', document:'', address:'', credit_limit:'0' });
  const [query, setQuery] = useState('');
  const filtered = customers.filter(c => `${c.name} ${c.phone} ${c.document}`.toLowerCase().includes(query.toLowerCase()));
  async function save(e) {
    e.preventDefault();
    if (!form.name.trim()) return alert('Coloca el nombre del cliente.');
    const payload = { ...form, credit_limit: asNum(form.credit_limit), status: 'Activo' };
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
  return <div className="page"><div className="hero compact-hero"><h1>📘 Inventario</h1><p>Vista compacta por categoría y stock.</p></div>{Object.entries(byCat).map(([cat, items]) => <section className="card compact-card inventory-block" key={cat}><h3>{cat}</h3>{items.map(p=><div className="list-row" key={p.id}><span>{p.code} · {p.name}<small>Stock mínimo {asNum(p.stock_min)}</small></span><strong className={asNum(p.stock) <= asNum(p.stock_min) ? 'danger-text' : ''}>Stock {asNum(p.stock)}</strong></div>)}</section>)}</div>;
}

function StockEntry({ products, reloadProducts }) {
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
    await supabase.from('stock_movements').insert({ product_id: selected.id, type: 'Entrada', qty: asNum(form.qty), note });
    await supabase.from('cash_movements').insert({ type: form.method === 'Crédito' ? 'Compra crédito' : 'Compra', payment_method: form.method, amount: asNum(form.paid || total), note: `Ingreso mercadería: ${selected.name}. ${note}` });
    alert(`Ingreso registrado. Nuevo stock: ${newStock}`);
    setForm({ product_id: selected.id, provider:'', qty:'1', cost:String(selected.cost || 0), method:'Efectivo', paid:'0', note:'' });
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
            <div className="list-row"><span>{selected.name}<small>{selected.code} · {selected.category}</small></span><strong>{money(selected.price)}</strong></div>
            <div className="list-row"><span>Stock actual</span><strong>{asNum(selected.stock)}</strong></div>
            <div className="list-row"><span>Stock después</span><strong>{asNum(selected.stock) + asNum(form.qty)}</strong></div>
          </> : <p className="muted">No hay productos.</p>}
        </section>
      </div>
    </div>
  );
}

function CashPage() {
  const { movements, reload } = useCashMovements();
  const [form, setForm] = useState({ type:'Ingreso', method:'Efectivo', amount:'0', note:'' });
  const today = todayISO();
  const todayMovs = movements.filter(m => String(m.created_at || '').slice(0,10) === today);
  const ingresos = todayMovs.filter(m => ['Ingreso','Apertura'].includes(m.type)).reduce((s,m)=>s+asNum(m.amount),0);
  const egresosList = todayMovs.filter(m => ['Egreso','Compra','Retiro'].includes(m.type));
  const egresos = egresosList.reduce((s,m)=>s+asNum(m.amount),0);
  const creditos = todayMovs.filter(m => String(m.type).includes('Crédito')).reduce((s,m)=>s+asNum(m.amount),0);
  const abonos = todayMovs.filter(m => String(m.note || '').toLowerCase().includes('abono')).reduce((s,m)=>s+asNum(m.amount),0);
  async function save(e) {
    e.preventDefault();
    if (asNum(form.amount) <= 0) return alert('El monto debe ser mayor a cero.');
    const { error } = await supabase.from('cash_movements').insert({ type: form.type, payment_method: form.method, amount: asNum(form.amount), note: form.note });
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


function Credits() {
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
      .or('status.eq.Crédito,payment_method.eq.Crédito')
      .order('created_at', { ascending: false });
    setSales(creditSales || []);
    const { data: pays, error: payError } = await supabase
      .from('credit_payments')
      .select('*')
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
    const payload = { sale_id: selected.id, customer_name: selected.customer_name || 'Cliente', amount, method: form.method, note: form.note };
    const { error } = await supabase.from('credit_payments').insert(payload);
    if (error) return alert('Falta ejecutar la actualización SQL de V01.4: tabla credit_payments.');
    await supabase.from('cash_movements').insert({ type: 'Ingreso', payment_method: form.method, amount, note: `Abono crédito B${selected.receipt_number}. ${form.note || ''}` });
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


function Reports({ products }) {
  const { sales } = useSales();
  const { movements } = useCashMovements();
  const [payments, setPayments] = useState([]);
  useEffect(() => { if (hasSupabaseConfig) supabase.from('credit_payments').select('*').then(({data}) => setPayments(data || [])); }, []);
  const totalSales = sales.reduce((s,v)=>s+asNum(v.total),0);
  const creditSales = sales.filter(s => s.status === 'Crédito' || s.payment_method === 'Crédito');
  const totalCredits = creditSales.reduce((s,v)=>s+asNum(v.total),0);
  const totalPaidCredits = payments.reduce((s,p)=>s+asNum(p.amount),0);
  const avg = sales.length ? totalSales / sales.length : 0;
  const byMethod = sales.reduce((acc, s) => { acc[s.payment_method || 'Sin método'] = (acc[s.payment_method || 'Sin método'] || 0) + asNum(s.total); return acc; }, {});
  const byCat = products.reduce((acc, p) => { acc[p.category || 'General'] = (acc[p.category || 'General'] || 0) + asNum(p.stock); return acc; }, {});
  const egresos = movements.filter(m => ['Egreso','Compra','Retiro'].includes(m.type)).reduce((s,m)=>s+asNum(m.amount),0);
  return <div className="page"><div className="hero compact-hero"><h1>📈 Reportes</h1><p>Ventas, caja, créditos, abonos e inventario en vista rápida.</p></div><div className="kpi-grid"><Kpi label="Total vendido" value={money(totalSales)} helper={`${sales.length} ventas`} /><Kpi label="Ticket promedio" value={money(avg)} helper="Promedio de venta" /><Kpi label="Crédito pendiente" value={money(Math.max(0,totalCredits-totalPaidCredits))} helper="Por cobrar" /><Kpi label="Egresos" value={money(egresos)} helper="Compras y salidas" /></div><div className="two-col"><section className="card compact-card"><h3>Ventas por método</h3>{Object.entries(byMethod).map(([k,v])=><div className="list-row" key={k}><span>{k}</span><strong>{money(v)}</strong></div>)}{!sales.length && <p className="muted">No hay ventas todavía.</p>}</section><section className="card compact-card"><h3>Stock por categoría</h3>{Object.entries(byCat).map(([k,v])=><div className="list-row" key={k}><span>{k}</span><strong>{v}</strong></div>)}</section></div><div className="two-col extra-row"><section className="card compact-card"><h3>Abonos recibidos</h3>{payments.slice(0,8).map(p=><div className="list-row" key={p.id}><span>{p.customer_name}<small>{fmtDate(p.created_at)} · {p.method}</small></span><strong>{money(p.amount)}</strong></div>)}{!payments.length && <p className="muted">No hay abonos registrados.</p>}</section><section className="card compact-card"><h3>Movimientos de caja</h3>{movements.slice(0,8).map(m=><div className="list-row" key={m.id}><span>{m.type} · {m.payment_method}<small>{m.note || 'Sin nota'}</small></span><strong>{money(m.amount)}</strong></div>)}</section></div></div>;
}


function AppShell({ session }) {
  const [current, setCurrent] = useState('ventas');
  const [open, setOpen] = useState(false);
  const { products, loading, reload } = useProducts();
  const { customers, reload: reloadCustomers } = useCustomers();
  const { sales } = useSales();
  const content = {
    panel: <Panel products={products}/>,
    ventas: <POS products={products} reloadProducts={reload} customers={customers}/>,
    productos: <Products products={products} reload={reload}/>,
    inventario: <Inventory products={products}/>,
    reportes: <Reports products={products}/>,
    creditos: <Credits/>,
    caja: <CashPage />,
    ingreso: <StockEntry products={products} reloadProducts={reload}/>,
    clientes: <Customers customers={customers} reload={reloadCustomers}/>,
  }[current];
  return <div className="app"><Sidebar current={current} setCurrent={setCurrent} open={open} setOpen={setOpen} session={session}/><main className="main"><Header setOpen={setOpen} current={current}/>{loading ? <div className="loader">Cargando...</div> : content}</main></div>;
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
