export interface Articulo {
    idArticulo: number;
    estado: boolean;
    nombre: string;
    marca: string | null;
    precio: number | null;
    stock: number | null;
    idCategoria: number | null;
}

export interface ArticuloCategoria {
    nombreArticulo: string;
    marca: string | null;
    nombreCategoria: string;
    precio: number | null;
    stock: number | null;
}
