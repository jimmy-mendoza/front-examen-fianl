import api from "../api/axios";
import type { ClienteFactura } from "../types/cliente";

// Consultar las facturas activas de clientes activos.
export const reporteClienteFactura = () => api.get<ClienteFactura[]>("/clientes/reporteClienteFactura");
