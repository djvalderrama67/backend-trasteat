# TrabajoGradoBack

## Descripción
Esta aplicación Django proporciona una API para la gestión de categorías, vehículos, bodegas, objetos y cálculos.

## Entidades Principales
1. **Categoria**: Representa diferentes categorías.
2. **Vehiculo**: Representa vehículos con propiedades como capacidad y volumen de carga.
3. **Bodega**: Entidad que representa bodegas (detalles no especificados en la revisión).
4. **Objeto**: Representa objetos que pueden estar asociados con categorías.
5. **Calculo**: Representa operaciones o cálculos realizados en la aplicación.

## Rutas API Disponibles
- **api/overview/**: Proporciona una visión general de las rutas API.
- **categoria/<str:nombre_categoria>/**: Devuelve objetos asociados con una categoría específica.
- **categoria/**: CRUD para categorías.
- **vehiculo/**: CRUD para vehículos.
- **bodega/**: CRUD para bodegas.
- **objeto/**: CRUD para objetos.
- **calculo/**: CRUD para cálculos

