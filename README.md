# Elizabeth Word Proccesor (ewop)

Programa en python que lee un listado de mensajes de texto dispuesto en la primera columna de 
una hoja excel y los procesa para obtener varios tipos de resultados, de acuerdo a los 
procesadores habilitados en su fichero de configuración.

Los procesadores incluidos actualmente son:

- **basic**: 
    + Conteo de repeticiones de palabras.
    
- **stanza (Librería Stanford NLP)**:
    + Extracción y conteo de frases relevantes.

## Requerimietos

### Librerías de python

- [stanza](https://github.com/stanfordnlp/stanza/)

## Modo de uso

```bash
./ewop.py -i my_data_file.xlsx
```

Para detalle de las opciones, ejecute:

```bash
./ewop.py --help
```
