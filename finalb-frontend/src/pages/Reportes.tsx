import { useRef, useState, type FormEvent } from "react";
import axios from "axios";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { listarArticulosActivosFiltro, reporteArticuloCategoria } from "../services/articuloServices";
import { reporteClienteFactura } from "../services/clienteServices";
import { logo } from "../assets/logo";

type TipoReporte = "articulos" | "categorias" | "facturas";
type Formato = "pdf" | "excel";
type Celda = string | number;

interface DatosReporte {
    titulo: string;
    archivo: string;
    columnas: string[];
    filas: Celda[][];
    anchos: number[];
    decimales: number[];
    filtro: string;
}

const opciones: { tipo: TipoReporte; numero: string; titulo: string; descripcion: string }[] = [
    { tipo: "articulos", numero: "01", titulo: "Artículos por nombre", descripcion: "Busca artículos activos por una parte de su nombre." },
    { tipo: "categorias", numero: "02", titulo: "Artículos con su categoría", descripcion: "Consulta marca, categoría, precio y existencias de cada artículo." },
    { tipo: "facturas", numero: "03", titulo: "Facturas por cliente", descripcion: "Consulta clientes, NIT, fechas y totales de facturas activas." }
];

function Reportes() {
    const dialogo = useRef<HTMLDialogElement>(null);
    const ocupado = useRef(false);
    const [dialogoAbierto, setDialogoAbierto] = useState(false);
    const [nombre, setNombre] = useState("");
    const [cargando, setCargando] = useState(false);
    const [mensaje, setMensaje] = useState("");
    const [error, setError] = useState("");
    const [reporte, setReporte] = useState<DatosReporte | null>(null);

    const formatearFecha = (valor: string | null) => {
        if (!valor) return "Sin fecha";
        const fecha = new Date(valor);
        return Number.isNaN(fecha.getTime()) ? "Fecha no válida" : fecha.toLocaleDateString("es-GT");
    };

    // Cada consulta utiliza su endpoint y conserva los números para exportar a Excel.
    const obtenerReporte = async (tipo: TipoReporte): Promise<DatosReporte> => {
        if (tipo === "articulos") {
            const respuesta = await listarArticulosActivosFiltro(nombre.trim());
            return {
                titulo: "Artículos por nombre", archivo: "articulos-por-nombre", columnas: ["Artículo", "Marca", "Precio", "Stock"],
                filas: respuesta.data.map((articulo) => [articulo.nombre, articulo.marca ?? "", articulo.precio ?? "", articulo.stock ?? ""]),
                anchos: [40, 30, 18, 16], decimales: [2], filtro: nombre.trim()
            };
        }
        if (tipo === "categorias") {
            const respuesta = await reporteArticuloCategoria();
            return {
                titulo: "Artículos con su categoría", archivo: "articulos-con-categoria", columnas: ["Artículo", "Marca", "Categoría", "Precio", "Stock"],
                filas: respuesta.data.map((articulo) => [articulo.nombreArticulo, articulo.marca ?? "", articulo.nombreCategoria, articulo.precio ?? "", articulo.stock ?? ""]),
                anchos: [40, 28, 30, 18, 16], decimales: [3], filtro: ""
            };
        }
        const respuesta = await reporteClienteFactura();
        return {
            titulo: "Facturas por cliente", archivo: "facturas-por-cliente", columnas: ["Cliente", "NIT", "Fecha", "Total"],
            filas: respuesta.data.map((cliente) => [[cliente.nombre, cliente.apellido].filter(Boolean).join(" "), cliente.nit ?? "", formatearFecha(cliente.fechaFactura), cliente.total ?? ""]),
            anchos: [48, 22, 22, 20], decimales: [3], filtro: ""
        };
    };

    const formatearCelda = (datos: DatosReporte, valor: Celda, indice: number) => {
        return datos.decimales.includes(indice) && typeof valor === "number" ? valor.toFixed(2) : String(valor);
    };

    // Repetir título, fecha y logo en cada página del PDF horizontal.
    const exportarPDF = (datos: DatosReporte) => {
        const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
        const fecha = new Date().toLocaleDateString("es-GT");
        const filtro = datos.filtro ? doc.splitTextToSize("Nombre contiene: " + datos.filtro, 230) as string[] : [];
        const inicioTabla = 45 + filtro.length * 5;
        autoTable(doc, {
            head: [datos.columnas],
            body: datos.filas.map((fila) => fila.map((valor, indice) => formatearCelda(datos, valor, indice))),
            startY: inicioTabla,
            margin: { top: inicioTabla, right: 14, bottom: 16, left: 14 },
            styles: { fontSize: 10, cellPadding: 3, overflow: "linebreak" },
            theme: "grid",
            headStyles: { fillColor: [30, 41, 59], fontStyle: "bold", cellPadding: 4 },
            alternateRowStyles: { fillColor: [255, 247, 237] },
            columnStyles: Object.fromEntries(datos.decimales.map((indice) => [indice, { halign: "right" as const }])),
            didDrawPage: () => {
                doc.setFillColor(15, 23, 42);
                doc.rect(0, 0, 297, 36, "F");
                doc.setFillColor(249, 115, 22);
                doc.rect(0, 36, 297, 1.2, "F");
                doc.setFillColor(255, 255, 255);
                doc.roundedRect(14, 8, 22, 22, 3, 3, "F");
                doc.addImage(logo, "PNG", 16, 10, 18, 18);
                doc.setFontSize(18);
                doc.setTextColor(255, 255, 255);
                doc.text(datos.titulo, 42, 18);
                doc.setFontSize(10);
                doc.setTextColor(203, 213, 225);
                doc.text("Fecha: " + fecha + " | Solo registros activos", 42, 27);
                doc.setTextColor(253, 186, 116);
                doc.text("FERRETERÍA", 283, 17, { align: "right" });
                doc.setTextColor(51, 65, 85);
                if (filtro.length) doc.text(filtro, 14, 44);
                doc.setDrawColor(203, 213, 225);
                doc.line(14, 197, 283, 197);
                doc.setFontSize(9);
                doc.text("Informes de ferretería", 14, 203);
                doc.text("Página " + doc.getCurrentPageInfo().pageNumber, 283, 203, { align: "right" });
            }
        });
        doc.save(datos.archivo + ".pdf");
    };

    // Dar formato a la hoja sin convertir los importes numéricos en texto.
    const exportarExcel = async (datos: DatosReporte) => {
        const libro = new ExcelJS.Workbook();
        const hoja = libro.addWorksheet("Reporte");
        hoja.mergeCells(1, 1, 1, datos.columnas.length);
        hoja.getCell("A1").value = datos.titulo;
        hoja.getCell("A1").font = { size: 20, bold: true, color: { argb: "FFFFFFFF" } };
        hoja.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
        hoja.getCell("A1").alignment = { vertical: "middle", indent: 1 };
        hoja.getRow(1).height = 42;
        hoja.addRow(["Fecha: " + new Date().toLocaleDateString("es-GT")]);
        hoja.addRow([datos.filtro ? "Nombre contiene: " + datos.filtro : "Solo registros activos"]);
        hoja.mergeCells(2, 1, 2, datos.columnas.length);
        hoja.mergeCells(3, 1, 3, datos.columnas.length);
        [2, 3].forEach((numero) => {
            hoja.getRow(numero).height = 23;
            hoja.getCell(numero, 1).font = { size: 11, color: { argb: "FF475569" } };
            hoja.getCell(numero, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF7ED" } };
            hoja.getCell(numero, 1).alignment = { vertical: "middle", indent: 1, wrapText: true };
        });
        hoja.addRow([]);
        hoja.mergeCells(4, 1, 4, datos.columnas.length);
        hoja.getCell("A4").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF97316" } };
        hoja.getRow(4).height = 5;
        const encabezado = hoja.addRow(datos.columnas);
        encabezado.height = 28;
        encabezado.eachCell((celda) => {
            celda.font = { bold: true, color: { argb: "FFFFFFFF" } };
            celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
            celda.alignment = { vertical: "middle" };
        });
        datos.filas.forEach((valores, indice) => {
            const fila = hoja.addRow(valores);
            fila.height = 25;
            fila.eachCell({ includeEmpty: true }, (celda, columna) => {
                celda.border = { bottom: { style: "thin", color: { argb: "FFE2E8F0" } } };
                celda.alignment = { vertical: "middle", wrapText: true };
                if (indice % 2 === 1) celda.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF7ED" } };
                if (datos.decimales.includes(columna - 1)) celda.numFmt = "0.00";
            });
        });
        datos.anchos.forEach((ancho, indice) => { hoja.getColumn(indice + 1).width = ancho; });
        hoja.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5, column: datos.columnas.length } };
        hoja.views = [{ state: "frozen", ySplit: 5 }];
        hoja.pageSetup = { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
        hoja.pageSetup.printTitlesRow = "1:5";
        const buffer = await libro.xlsx.writeBuffer();
        const blob = new Blob([new Uint8Array(buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const enlace = document.createElement("a");
        enlace.href = url;
        enlace.download = datos.archivo + ".xlsx";
        document.body.appendChild(enlace);
        enlace.click();
        enlace.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    };

    // Consultar al exportar para que el archivo corresponda al filtro actual.
    const generarReporte = async (tipo: TipoReporte, formato: Formato) => {
        if (ocupado.current) return;
        if (tipo === "articulos" && !nombre.trim()) {
            setError("Escribe una parte del nombre del artículo.");
            return;
        }
        ocupado.current = true;
        setCargando(true);
        setMensaje("");
        setError("");
        try {
            const datos = await obtenerReporte(tipo);
            setReporte(datos);
            if (!datos.filas.length) {
                setMensaje("No se encontraron registros activos para este reporte.");
                return;
            }
            if (formato === "pdf") exportarPDF(datos);
            else await exportarExcel(datos);
            setMensaje("Reporte exportado a " + (formato === "pdf" ? "PDF" : "Excel") + ": " + datos.filas.length + " registros.");
        } catch (error) {
            const detalle = axios.isAxiosError(error) ? error.response?.data?.mensaje ?? error.message : error instanceof Error ? error.message : "Error inesperado";
            setError("No se pudo generar el reporte: " + detalle);
        } finally {
            ocupado.current = false;
            setCargando(false);
        }
    };

    const abrirFiltro = () => {
        setMensaje("");
        setError("");
        setDialogoAbierto(true);
        dialogo.current?.showModal();
    };

    const handleSubmit = (evento: FormEvent<HTMLFormElement>) => {
        evento.preventDefault();
        void generarReporte("articulos", "pdf");
    };

    const clasePDF = "rounded-lg bg-slate-900 px-5 py-3 text-sm font-bold text-white transition hover:bg-slate-700 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-orange-600 disabled:opacity-50";
    const claseExcel = "rounded-lg border border-orange-300 bg-orange-50 px-5 py-3 text-sm font-bold text-orange-900 transition hover:bg-orange-100 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-orange-600 disabled:opacity-50";

    return (
        <main className="min-h-screen bg-stone-50">
            <header className="border-b-4 border-orange-500 bg-slate-900 px-5 py-7 text-white sm:px-9">
                <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-5">
                    <div className="flex items-center gap-4">
                        <img src={logo} alt="Logo de la ferretería" className="h-14 w-14 rounded-lg bg-white object-contain p-2" />
                        <div>
                            <p className="mb-1 text-xs font-bold uppercase tracking-[0.2em] text-orange-300">Administración de ferretería</p>
                            <p className="text-xl font-bold tracking-tight">Consultas y reportes</p>
                        </div>
                    </div>
                    <span className="flex items-center gap-2 rounded-lg border border-slate-600 bg-slate-800 px-4 py-2 text-xs font-semibold text-slate-200"><span className="h-2 w-2 rounded-full bg-orange-400" />Solo registros activos</span>
                </div>
            </header>
            <div className="mx-auto grid max-w-7xl gap-8 px-5 py-8 sm:px-9 lg:grid-cols-[220px_1fr] lg:py-11">
                <aside className="self-start rounded-xl border border-stone-200 bg-white p-5 lg:sticky lg:top-6">
                    <p className="mb-4 text-xs font-bold uppercase tracking-widest text-slate-400">Biblioteca de informes</p>
                    <nav aria-label="Secciones de reportes" className="flex flex-col gap-2">
                        <a href="#reportes" className="rounded-lg border-l-4 border-orange-500 bg-orange-50 px-4 py-3 text-sm font-bold text-orange-900">Reportes disponibles</a>
                        <a href="#ultima-consulta" className="rounded-lg px-4 py-3 text-sm text-slate-600 transition hover:bg-stone-100">Última consulta</a>
                    </nav>
                    <div className="mt-6 border-t border-stone-200 pt-5">
                        <p className="mb-2 text-sm font-bold text-slate-800">Formatos de descarga</p>
                        <p className="text-xs leading-6 text-slate-500">PDF para imprimir y compartir.<br />Excel para consultar y organizar.</p>
                    </div>
                </aside>
                <div className="min-w-0">
                    <section id="reportes" aria-labelledby="titulo-reportes">
                        <div className="mb-7 flex flex-wrap items-end justify-between gap-4">
                            <div>
                                <p className="mb-2 text-xs font-bold uppercase tracking-widest text-orange-700">Inventario y facturación</p>
                                <h1 id="titulo-reportes" className="text-3xl font-bold tracking-tight text-slate-900 sm:text-4xl">Informes de ferretería</h1>
                                <p className="mt-3 max-w-xl text-sm leading-6 text-slate-500">Elige la información que necesitas y selecciona un formato de descarga.</p>
                            </div>
                            <span className="rounded-lg border border-stone-200 bg-white px-3 py-2 text-xs font-bold text-slate-500">03 reportes</span>
                        </div>
                        <div aria-live="polite" className="mb-5 space-y-3">
                            {cargando && <p className="rounded-lg border-l-4 border-orange-500 bg-orange-50 px-4 py-3 text-sm text-orange-900">Preparando tu reporte…</p>}
                            {!dialogoAbierto && error && <p role="alert" className="rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{error}</p>}
                            {!dialogoAbierto && mensaje && <p className="rounded-lg border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{mensaje}</p>}
                        </div>
                        <div className="space-y-4">
                            {opciones.map((opcion) => (
                                <article key={opcion.tipo} className="animate-fade-in flex flex-col gap-5 rounded-xl border border-stone-200 bg-white p-5 transition duration-200 hover:border-orange-300 hover:shadow-sm motion-reduce:animate-none sm:p-6 xl:flex-row xl:items-center">
                                    <span className="flex h-12 w-12 shrink-0 items-center justify-center rounded-lg bg-slate-100 font-mono text-lg font-bold text-slate-500">{opcion.numero}</span>
                                    <div className="flex-1">
                                        <h2 className="mb-2 text-lg font-bold text-slate-900">{opcion.titulo}</h2>
                                        <p className="max-w-md text-sm leading-6 text-slate-500">{opcion.descripcion}</p>
                                    </div>
                                    {opcion.tipo === "articulos" ? (
                                        <button type="button" onClick={abrirFiltro} disabled={cargando} className={clasePDF}>Buscar artículos →</button>
                                    ) : (
                                        <div className="flex shrink-0 flex-wrap gap-2">
                                            <button type="button" onClick={() => void generarReporte(opcion.tipo, "pdf")} disabled={cargando} className={clasePDF}>Exportar a PDF</button>
                                            <button type="button" onClick={() => void generarReporte(opcion.tipo, "excel")} disabled={cargando} className={claseExcel}>Exportar a Excel</button>
                                        </div>
                                    )}
                                </article>
                            ))}
                        </div>
                    </section>
                    <section id="ultima-consulta" aria-labelledby="titulo-consulta" className="mt-8 overflow-hidden rounded-xl border border-stone-200 bg-white">
                        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-stone-200 bg-slate-100 px-6 py-5">
                            <div>
                                <p className="mb-1 text-xs font-bold uppercase tracking-widest text-slate-500">Última consulta</p>
                                <h2 id="titulo-consulta" className="text-lg font-bold text-slate-900">{reporte?.titulo ?? "Vista previa del reporte"}</h2>
                                {reporte?.filtro && <p className="mt-1 text-sm text-slate-500">Nombre contiene: {reporte.filtro}</p>}
                            </div>
                            {reporte && <span className="rounded-md bg-slate-900 px-3 py-1.5 text-xs font-bold text-white">{reporte.filas.length} registros</span>}
                        </div>
                        {reporte ? (
                            <div className="overflow-x-auto">
                                <table className="w-full text-left text-sm">
                                    <thead className="bg-slate-800 text-white">
                                        <tr>{reporte.columnas.map((columna, indice) => <th key={columna} scope="col" className={"px-5 py-4 font-semibold " + (reporte.decimales.includes(indice) ? "text-right" : "")}>{columna}</th>)}</tr>
                                    </thead>
                                    <tbody className="divide-y divide-stone-100">
                                        {reporte.filas.map((fila, indice) => (
                                            <tr key={indice} className="even:bg-orange-50/50 hover:bg-orange-100/60">
                                                {fila.map((valor, columna) => <td key={columna} className={"px-5 py-4 tabular-nums " + (reporte.decimales.includes(columna) ? "text-right font-semibold" : "")}>{formatearCelda(reporte, valor, columna)}</td>)}
                                            </tr>
                                        ))}
                                        {!reporte.filas.length && <tr><td colSpan={reporte.columnas.length} className="px-6 py-10 text-center text-slate-500">No se encontraron registros activos.</td></tr>}
                                    </tbody>
                                </table>
                            </div>
                        ) : (
                            <div className="px-6 py-12 text-center">
                                <span aria-hidden="true" className="mx-auto mb-4 flex h-12 w-12 items-center justify-center rounded-lg border border-dashed border-stone-300 text-2xl text-slate-400">≡</span>
                                <p className="text-sm font-semibold text-slate-600">Todavía no has consultado un reporte</p>
                                <p className="mt-2 text-xs text-slate-400">Los resultados aparecerán aquí después de exportar.</p>
                            </div>
                        )}
                    </section>
                    <footer className="mt-6 flex flex-wrap justify-between gap-2 border-t border-stone-200 pt-5 text-xs text-slate-400"><span>Ferretería · Consultas y reportes</span><span>PDF horizontal / Excel</span></footer>
                </div>
            </div>
            <dialog ref={dialogo} aria-labelledby="titulo-filtro" onClose={() => setDialogoAbierto(false)} onCancel={(evento) => { if (cargando) evento.preventDefault(); }} className="fixed inset-0 m-auto w-[calc(100%-2rem)] max-w-lg overflow-hidden rounded-xl border-0 bg-white p-0 shadow-2xl">
                <form onSubmit={handleSubmit}>
                    <div className="flex items-start justify-between gap-4 border-b-4 border-orange-500 bg-slate-900 px-7 py-6 text-white">
                        <div>
                            <p className="mb-2 text-xs font-bold uppercase tracking-widest text-orange-300">Filtro de inventario</p>
                            <h2 id="titulo-filtro" className="text-2xl font-bold">Artículos por nombre</h2>
                        </div>
                        <button type="button" aria-label="Cerrar diálogo" onClick={() => dialogo.current?.close()} disabled={cargando} className="rounded-md px-2 py-1 text-xl text-slate-300 hover:bg-slate-700 disabled:opacity-50">×</button>
                    </div>
                    <div className="p-7">
                        <label htmlFor="nombre" className="mb-2 block text-sm font-bold text-slate-700">¿Qué artículo necesitas?</label>
                        <input id="nombre" name="nombre" type="text" value={nombre} onChange={(evento) => setNombre(evento.target.value)} placeholder="Escribe un nombre, por ejemplo: martillo" required maxLength={65} disabled={cargando} className="w-full rounded-lg border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none transition focus:border-orange-500 focus:bg-white focus:ring-4 focus:ring-orange-50" />
                        <p className="mt-3 text-xs leading-5 text-slate-500">Puedes escribir solo una parte del nombre. La búsqueda no distingue mayúsculas y minúsculas.</p>
                        <div aria-live="polite" className="mt-4 text-sm">
                            {cargando && <p className="text-orange-800">Preparando el reporte…</p>}
                            {error && <p role="alert" className="text-red-700">{error}</p>}
                            {mensaje && <p className="text-emerald-800">{mensaje}</p>}
                        </div>
                        <div className="mt-6 flex flex-wrap gap-3 border-t border-stone-200 pt-5">
                            <button type="submit" disabled={cargando || !nombre.trim()} className={clasePDF}>Exportar a PDF</button>
                            <button type="button" onClick={() => void generarReporte("articulos", "excel")} disabled={cargando || !nombre.trim()} className={claseExcel}>Exportar a Excel</button>
                        </div>
                    </div>
                </form>
            </dialog>
        </main>
    );
}

export default Reportes;
