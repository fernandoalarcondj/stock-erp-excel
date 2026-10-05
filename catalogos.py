"""
CATÁLOGOS DE VALORES CONOCIDOS
==============================
Si el ERP agrega un producto, calidad o alistamiento nuevo, o escribe uno de
otra forma, se agrega acá y no hace falta tocar el resto del código.

Formato:  "Nombre que aparecerá en el Excel": ["TEXTO COMO SALE EN EL PDF", ...]
Los textos del PDF van en MAYÚSCULAS. El orden dentro de la lista no importa
(se prueban primero los más largos). Todo lo que no coincida queda en blanco.
"""

CALIDADES = {
    "Standart": ["STANDART", "STANDAR", "STANDA", "STAND", "STANDARD", "ESTANDAR"],
    "Bifaz ST": ["BIFAZ STANDART", "BIFAZ STANDAR", "BIFAZ ST", "BIFAZ S", "BIFAZ"],
    "Hojas Standart": [
        "HOJAS STANDART", "HOJAS STANDAR", "HOJAS STAND", "HOJAS ST", "HOJAS S",
        "HOJA STANDART", "HOJA STANDAR", "HOJA ST", "HOJA S",
    ],
    "Hojas Bifaz": [
        "HOJAS BIFAZ ST", "HOJAS BIFAZ S", "HOJAS BIFAZ", "HOJAS B",
        "HOJA BIFAZ", "HOJA B",
    ],
    "Rebobinado": ["REBOBINADO", "REBOBINAD", "REBOBINA", "REBOBIN", "REBOB"],
}

PRODUCTOS = {
    "Liner Per": ["LINER PERLADO", "LINER PERL", "LINER PER"],
    "Liner Blanco": ["LINER BLANCO", "LINER BLAN", "LINER BLA", "LINER B"],
    "Onda C": ["ONDA C"],
    "Onda Liner": ["ONDA LINER", "ONDA LINE", "ONDA LIN"],
    "Cartulina Blanca": ["CARTULINA BLANCA", "CARTULINA BLAN", "CART BLANCA", "CART BLAN", "C BLANCA", "C BLAN"],
    "Cartulina Gris": ["CARTULINA GRIS", "CART GRIS", "C GRIS"],
    "Covering": ["COVERING", "COVER"],
}

# El ERP a veces imprime solo "CARTULINA" (sin color). En ese caso se deduce el
# color por la calidad. Esto surge del resumen del propio reporte Stock Puro:
# Cartulina Bifaz ST -> columna C.BLANCA ; Cartulina Standart -> columna C.GRIS.
# Si la calidad no está en este diccionario, el producto queda en blanco.
CARTULINA_GENERICA = ["CARTULINA", "CARTUL", "CART"]
CARTULINA_SEGUN_CALIDAD = {
    "Bifaz ST": "Cartulina Blanca",
    "Hojas Bifaz": "Cartulina Blanca",
    "Standart": "Cartulina Gris",
    "Hojas Standart": "Cartulina Gris",
}

ALISTAMIENTOS = {
    "Bobina": ["BOBINA", "BOBI"],
    "Torta": ["TORTA", "TORT"],
    "Bancal": ["BANCAL", "BANC"],
    "Tubo": ["TUBO"],
}

OBSERVACIONES = {
    "Encolado": ["ENCOLADO", "ENCOLAD"],
}
