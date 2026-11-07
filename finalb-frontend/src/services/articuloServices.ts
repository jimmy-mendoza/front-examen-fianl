import api from "../api/axios";
import type { Articulo, ArticuloCategoria } from "../types/articulo";

// El filtro se envía como parámetro para conservar espacios y caracteres especiales.
export const listarArticulosActivosFiltro = (nombre: string) => api.get<Articulo[]>("/articulos/mostrarActivosFiltro", { params: { nombre } });
export const reporteArticuloCategoria = () => api.get<ArticuloCategoria[]>("/articulos/reporteArticuloCategoria");
