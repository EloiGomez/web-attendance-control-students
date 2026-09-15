[English](README.md) | **Català** | [Español](README.es.md)

# Control d'Assistència

Un sistema lleuger de control d'assistència escolar fet amb **Google Apps Script**, creat per substituir un Excel lent i incòmode amb desenes de pestanyes que feia servir tota una escola de primària per anotar les absències diàries dels alumnes.

> Construït com a projecte de programació en parella amb IA amb [Claude](https://claude.com/claude-code): jo vaig dirigir els requisits, les decisions de disseny i les proves (incloent-hi detectar un error de caché, un error de fallada silenciosa i un problema de rendiment amb molta càrrega de dades), mentre que Claude s'ha encarregat de la implementació.

## El problema

El flux de treball original era un únic fitxer Excel amb una graella enorme per cada classe (una fila per alumne, una columna per cada dia lectiu de l'any — més o menys 240 files × 390 columnes per classe), duplicada en desenes de pestanyes per a tota l'escola. Els mestres havien d'escriure codis d'una lletra a mà a cada cel·la, el fitxer trigava molt a obrir-se i editar-se, i no hi havia cap manera automàtica de veure el percentatge d'absències d'un alumne durant el curs.

## La solució

Una petita aplicació web servida directament des de Google Apps Script, amb una Google Sheet al darrere que només fa de base de dades (els mestres no l'obren mai directament):

- **Passar llista per classe i dia**: tries una classe i una data, veus tots els alumnes en una graella de caselles (matí / tarda / justificada / retard), i guardes tot el dia d'un cop.
- **Resum d'absències en temps real**: percentatge d'assistència automàtic per alumne (ponderat — faltar un matí no compta igual que faltar una tarda, seguint l'horari real de l'escola), desglossat entre falta justificada/no justificada, amb un nivell d'alerta progressiu per colors, un desglossament per dia de la setmana, un cercador i columnes ordenables.
- **Control d'accés**: restringit a una llista explícita d'emails de professorat autoritzat, comprovada al servidor a cada petició — no només un enllaç difícil d'endevinar.
- **Eines d'administració**: una eina d'un sol clic per "promocionar totes les classes" a final de curs (amb un full de mapeig de classes editable i una casella de repetidor per alumne), a més de generadors de dades de prova per fer proves de càrrega amb centenars d'alumnes.
- **Rendiment**: el backend agrupa totes les lectures de la Sheet, reutilitza un únic objecte `Spreadsheet` per execució en comptes de tornar-lo a demanar repetidament, i guarda en caché el resum calculat (comprimit amb gzip perquè càpiga dins el límit de mida del caché d'Apps Script), de manera que tornar a consultar-lo és gairebé instantani fins que les dades canvien de veritat.
- Mode fosc automàtic (`prefers-color-scheme`), un disseny adaptat a mòbil amb capçaleres/columnes fixes en fer scroll, i cap dependència externa — HTML/CSS/JS pur servit a través de `HtmlService`.

## Tecnologia

- **Backend**: Google Apps Script (JavaScript, motor V8) — `SpreadsheetApp`, `HtmlService`, `CacheService`, `PropertiesService`.
- **Frontend**: HTML/CSS/JS pur, sense frameworks ni pas de compilació — es comunica amb el backend via `google.script.run`.
- **Emmagatzematge de dades**: una Google Sheet, feta servir com una base de dades estructurada senzilla, no com a interfície.

## Estat

Prototip funcional, provat amb dades sintètiques (centenars d'alumnes, milers de registres d'assistència). No conté cap dada real d'alumnes — el full `Students` porta només noms d'exemple.

## Instal·lació i desplegament

No cal cap entorn local ni pas de compilació — tot funciona dins la infraestructura de Google.

1. Crea una nova Google Sheet en blanc a [sheets.google.com](https://sheets.google.com).
2. Obre **Extensions → Apps Script**. Això crea un projecte d'Apps Script vinculat a aquesta Sheet.
3. Al fitxer `Code.gs` per defecte, esborra el contingut d'exemple i enganxa-hi el [`Code.gs`](Code.gs) d'aquest repositori.
4. Afegeix un fitxer nou de tipus **HTML** (la `+` al costat d'"Archivos"), anomena'l exactament `Index` (sense extensió), i enganxa-hi l'[`Index.html`](Index.html) d'aquest repositori.
5. Desa el projecte.
6. Al desplegable de funcions de dalt, selecciona **`setup`** i pulsa **Executar**. La primera vegada et demanarà autoritzar l'script — accepta-ho. Això crea els fulls `Config`, `Students`, `Teachers`, `Records` i `Promotion` amb dades d'exemple i valors per defecte raonables.
7. (Opcional) Executa **`generateTestStudents`** i **`generateTestRecords`** per carregar uns centenars d'alumnes i registres d'assistència sintètics, útil per provar el resum i l'ordenació/cercador sense escriure res a mà.
8. Afegeix els emails de qui hagi de tenir accés al full **`Teachers`** — si aquest full està buit, l'aplicació deixa entrar a tothom, així que aquest pas importa abans de compartir l'enllaç.
9. Ves a **Implementar → Nova implementació**, tria el tipus **Aplicació web**. Configura:
   - **Executar com**: *Usuari que accedeix a l'aplicació web* — així la identitat de Google de cada visitant és la que fa servir de veritat el control d'accés (i la columna d'auditoria "Updated By").
   - **Qui té accés**: *Qualsevol usuari amb compte de Google* (o *Qualsevol usuari dins de [domini]*, si es desplega des d'una organització de Google Workspace).
10. Pulsa **Implementar**, i obre la URL `.../exec` resultant. La primera vegada que cada usuari l'obri, Google li demanarà autoritzar l'aplicació amb el seu propi compte — és normal, ja que s'executa com cada visitant, no com el desenvolupador.

Per publicar un canvi de codi més endavant: edita els fitxers, desa, i després **Implementar → Gestionar implementacions → ✏️ → Versió: Nova versió → Implementar** — la mateixa URL `.../exec` segueix funcionant, només comença a servir el codi actualitzat.

## Fitxers

- [`Code.gs`](Code.gs) — lògica del servidor (rutes, control d'accés, accés a dades, lògica de negoci, eines d'administració).
- [`Index.html`](Index.html) — el frontend d'una sola pàgina.
