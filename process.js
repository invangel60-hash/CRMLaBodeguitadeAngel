const xlsx = require('xlsx');

const workbook = xlsx.readFile('ListadodeProductos.xlsx');
const sheetName = workbook.SheetNames[0];
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

// Helper to remove extra spaces
const cleanStr = str => str ? str.trim().replace(/\s+/g, ' ') : '';

// Helper to remove measurements
const removeMeasurements = name => {
  return name.replace(/\b(\d+([.,]\d+)?)\s*(ml|lts|litro|litros|g|gr|grs|kg|cc|oz|lb|cm)\b/gi, '')
             .replace(/\b(\d+)\s*(mil|pack|und|unidades|paquetes)\b/gi, '')
             .replace(/[^a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\s]/g, '')
             .replace(/\s+/g, ' ').trim();
};

const substitutions = {
  'aceite': { syn: ['aceites', 'grasa', 'cocina', 'freir', 'viveres', 'despensa', 'vegetal', 'maiz', 'soya'], brands: ['Amacorp', 'Anguie', 'Chef', 'Coamo', 'Concordia', 'Ideal', 'La Pampa', 'Mazeite', 'Primor', 'Vatel', 'Soya'] },
  'harina de maiz': { syn: ['harina', 'arepa', 'masa', 'viveres', 'despensa', 'maiz'], brands: ['Pan', 'Juana', 'Doña Arepa', 'Demasa', 'Lucharepa'] },
  'harina de trigo': { syn: ['harina', 'trigo', 'panaderia', 'reposteria', 'torta', 'viveres', 'despensa'], brands: ['Robin Hood', 'Blanca', 'Leudante', 'Todo uso'] },
  'arroz': { syn: ['granos', 'carbohidrato', 'viveres', 'despensa', 'comida'], brands: ['Mary', 'Primor', 'Santa Ana', 'Corina', 'Marquesa'] },
  'pasta': { syn: ['fideos', 'spaghetti', 'macarrones', 'carbohidrato', 'viveres', 'despensa', 'comida'], brands: ['Mary', 'Primor', 'Capri', 'Sindoni', 'Ronco', 'Allegri'] },
  'azucar': { syn: ['endulzante', 'dulce', 'viveres', 'despensa', 'reposteria'], brands: ['Montalban', 'Rio Turbio', 'Providencia'] },
  'cafe': { syn: ['desayuno', 'bebida caliente', 'viveres', 'despensa'], brands: ['Fama de America', 'Madrid', 'Flor de Patria', 'San Antonio', 'Amanecer', 'El Peñon'] },
  'mantequilla': { syn: ['margarina', 'grasa', 'untar', 'desayuno', 'viveres', 'lacteos'], brands: ['Mavesa', 'Nelly', 'Chiffon', 'Deline'] },
  'margarina': { syn: ['mantequilla', 'grasa', 'untar', 'desayuno', 'viveres', 'lacteos'], brands: ['Mavesa', 'Nelly', 'Chiffon', 'Deline'] },
  'leche': { syn: ['lacteos', 'bebida', 'desayuno', 'viveres', 'polvo', 'liquida'], brands: ['Campestre', 'Zulia', 'Los Andes', 'La Campiña', 'Parmalat', 'Tole'] },
  'queso': { syn: ['lacteos', 'charcuteria', 'desayuno', 'cena', 'untar', 'rebanado'], brands: ['Los Andes', 'Torondoy', 'Pais', 'Zulia'] },
  'jamon': { syn: ['embutidos', 'charcuteria', 'desayuno', 'cena', 'rebanado', 'fiambre'], brands: ['Plumrose', 'Hermógenes', 'Oscar Mayer', 'Alimenta'] },
  'salchicha': { syn: ['embutidos', 'charcuteria', 'perro caliente', 'fiambre'], brands: ['Plumrose', 'Hermógenes', 'Oscar Mayer', 'Alimenta'] },
  'mayonesa': { syn: ['salsa', 'aderezo', 'viveres', 'despensa', 'untar'], brands: ['Mavesa', 'Kraft', 'Alfonzo Rivas', 'Nelly'] },
  'salsa de tomate': { syn: ['ketchup', 'salsa', 'aderezo', 'viveres', 'despensa'], brands: ['Pampero', 'Heinz', 'Mavesa', 'Fritz'] },
  'refresco': { syn: ['bebida', 'soda', 'gaseosa', 'cola', 'frio'], brands: ['Coca Cola', 'Pepsi', 'Glup', 'Dumbo', 'Golden', 'Hit', '7Up', 'Chinotto'] },
  'malta': { syn: ['bebida', 'refresco', 'cebada', 'energia', 'frio'], brands: ['Polar', 'Regional', 'Caracas'] },
  'jugo': { syn: ['bebida', 'fruta', 'nectar', 'frio'], brands: ['Yukery', 'Natulac', 'Frica', 'Santal', 'Los Andes'] },
  'galleta': { syn: ['snack', 'chucheria', 'dulce', 'merienda', 'postre'], brands: ['Oreo', 'Club Social', 'Maria', 'Susy', 'Cocosette', 'Samba', 'Reinitas'] },
  'chocolate': { syn: ['snack', 'chucheria', 'dulce', 'cacao', 'merienda', 'postre'], brands: ['Savoy', 'Carré', 'Galak', 'Cri-Cri', 'Ping Pong', 'Bolero', 'Corona'] },
  'caramelo': { syn: ['snack', 'chucheria', 'dulce', 'golosina'], brands: ['Locatel', 'Alpina', 'Trident', 'Halls', 'Certts'] },
  'snack': { syn: ['chucheria', 'salado', 'pasapalo', 'saladito', 'merienda'], brands: ['Pepito', 'Cheese Tris', 'Doritos', 'Ruffles', 'Lays', 'Natuchips', 'Jumbo Riko', 'Puff', 'Jack'] },
  'pepito': { syn: ['chucheria', 'salado', 'pasapalo', 'saladito', 'snack', 'queso', 'maiz'], brands: ['Cheese Tris', 'Doritos', 'Jumbo Riko', 'Puff', 'Jack'] },
  'cheese tris': { syn: ['chucheria', 'salado', 'pasapalo', 'saladito', 'snack', 'queso', 'maiz', 'pepito'], brands: ['Pepito', 'Doritos', 'Jumbo Riko', 'Puff', 'Jack'] },
  'doritos': { syn: ['chucheria', 'salado', 'pasapalo', 'saladito', 'snack', 'queso', 'maiz', 'tortilla'], brands: ['Cheese Tris', 'Pepito', 'Jumbo Riko', 'Puff', 'Jack'] },
  'jumbo riko': { syn: ['chucheria', 'salado', 'pasapalo', 'saladito', 'snack', 'queso', 'maiz', 'pepito'], brands: ['Cheese Tris', 'Doritos', 'Pepito', 'Puff', 'Jack'] },
  'chupeta': { syn: ['snack', 'chucheria', 'dulce', 'golosina', 'caramelo'], brands: ['Bon Bon Bum', 'Kool Aid', 'Tico Tico'] },
  'desodorante': { syn: ['aseo personal', 'higiene', 'cuidado', 'antitranspirante'], brands: ['Mum', 'Rexona', 'Dove', 'Speed Stick', 'Lady Speed Stick', 'Axe', 'Old Spice'] },
  'shampoo': { syn: ['aseo personal', 'higiene', 'cuidado', 'cabello', 'pelo'], brands: ['Head & Shoulders', 'Pantene', 'Drene', 'Every Night'] },
  'jabon': { syn: ['aseo personal', 'higiene', 'cuidado', 'piel', 'baño', 'limpieza', 'ropa'], brands: ['Protex', 'Dove', 'Palmolive', 'Moncler', 'Las Llaves', 'Ariel', 'Ace'] },
  'papel higienico': { syn: ['aseo personal', 'higiene', 'baño', 'rollo'], brands: ['Rosal', 'Scott', 'Sutil', 'Suave'] },
  'crema dental': { syn: ['aseo personal', 'higiene', 'dientes', 'pasta', 'bucal'], brands: ['Colgate', 'Crest', 'Oral-B'] },
  'toallas sanitarias': { syn: ['aseo personal', 'higiene', 'mujer', 'intimo'], brands: ['Always', 'Kotex', 'Stayfree', 'Nosotras'] },
  'detergente': { syn: ['limpieza', 'ropa', 'lavar', 'polvo', 'liquido'], brands: ['Ariel', 'Ace', 'Las Llaves', 'Nevex'] },
  'limpiador': { syn: ['limpieza', 'hogar', 'piso', 'desinfectante'], brands: ['Mister Musculo', 'Mistolin', 'Clorox', 'Nevex'] },
  'cloro': { syn: ['limpieza', 'hogar', 'blanqueador', 'desinfectante'], brands: ['Clorox', 'Nevex'] },
  'cigarros': { syn: ['tabaco', 'fumar', 'cigarro', 'cigarrillo', 'nicotina'], brands: ['Belmont', 'Consul', 'Viceroy', 'L&M', 'Marlboro', 'Chesterfield', 'Lucky Strike'] },
  'cerveza': { syn: ['licor', 'bebida', 'alcohol', 'fria'], brands: ['Polar', 'Regional', 'Zulia', 'Solera'] },
  'licor': { syn: ['bebida', 'alcohol', 'ron', 'vodka', 'whisky'], brands: ['Cacique', 'Santa Teresa', 'Pampero', 'Gordons', 'Bajo Cero'] },
  'agua': { syn: ['bebida', 'hidratacion', 'mineral', 'potable'], brands: ['Minalba', 'Nevada', 'Canaima'] }
};

const fallbacksByClassification = {
  'Abarrotes': { syn: ['viveres', 'despensa', 'mercado', 'alimento'] },
  'Aseo Personal y Limpieza': { syn: ['aseo', 'limpieza', 'higiene', 'cuidado', 'hogar'] },
  'Chucheria': { syn: ['snack', 'dulce', 'salado', 'golosina', 'pasapalo', 'saladito', 'merienda'] },
  'Cigarros': { syn: ['tabaco', 'fumar', 'vicio', 'cigarro', 'cigarrillo', 'nicotina'], brands: ['Belmont', 'Consul', 'Viceroy', 'L&M', 'Marlboro', 'Chesterfield', 'Lucky Strike'] },
  'Proteina': { syn: ['carne', 'pollo', 'pescado', 'cerdo', 'embutidos', 'comida', 'alimento'] },
  'Refrescos': { syn: ['bebida', 'liquido', 'soda', 'gaseosa', 'cola'] },
  'Quincalleria': { syn: ['utiles', 'varios', 'herramientas', 'hogar'] },
  'General': { syn: ['varios'] }
};

data.forEach(row => {
  const clasificacion = cleanStr(row.Clasificacion);
  const productoOriginal = cleanStr(row.Producto);
  
  let tags = new Set();
  
  // 3. Limpieza de Nombre
  const cleanName = removeMeasurements(productoOriginal);
  tags.add(cleanName);
  
  // Split into words for tag matching
  const words = cleanName.toLowerCase().split(/\s+/);
  words.forEach(w => {
    if (w.length > 2) tags.add(w);
  });
  
  // 1 & 2. Relación Cruzada / Sustitutos / Sinónimos
  const lowerName = cleanName.toLowerCase();
  
  let foundMatch = false;
  
  for (const [key, val] of Object.entries(substitutions)) {
    if (lowerName.includes(key)) {
      foundMatch = true;
      if (val.syn) val.syn.forEach(s => tags.add(s));
      if (val.brands) val.brands.forEach(b => tags.add(b));
    }
  }
  
  // Use fallback category tags if no specific match
  if (!foundMatch && fallbacksByClassification[clasificacion]) {
      const fallback = fallbacksByClassification[clasificacion];
      if (fallback.syn) fallback.syn.forEach(s => tags.add(s));
      if (fallback.brands) fallback.brands.forEach(b => tags.add(b));
  }
  
  // Extra specific hacks based on prompt rules:
  // "cheese tris", "Jumbo Riko", "Pepito", "Doritos", "Puff"
  if (['cheese', 'tris', 'jumbo', 'riko', 'pepito', 'doritos', 'puff', 'snack', 'tostitos', 'lays'].some(v => lowerName.includes(v))) {
      const syn = ['chucheria', 'salado', 'pasapalo', 'saladito', 'snack', 'queso', 'maiz', 'viveres'];
      const brands = ['Cheese Tris', 'Pepito', 'Doritos', 'Jumbo Riko', 'Puff', 'Jack', 'Natuchips', 'Lays', 'Ruffles', 'De Todito'];
      syn.forEach(s => tags.add(s));
      brands.forEach(b => tags.add(b));
  }

  // Convert Set to comma-separated string
  row.Informacion = Array.from(tags).join(', ');
  
  // Cleanup original keys
  row.Clasificacion = row.Clasificacion;
  row.IdProducto = row.IdProducto;
  row.Producto = row.Producto;
});

const newWorkbook = xlsx.utils.book_new();
const newWorksheet = xlsx.utils.json_to_sheet(data);
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
xlsx.writeFile(newWorkbook, 'ListadodeProductos_Actualizado.xlsx');

console.log('Done processing. File saved as ListadodeProductos_Actualizado.xlsx');
