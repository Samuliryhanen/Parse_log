Tämä on pieni projekti tehtynä Genretech Oy:lle. Scriptin ideana on käsitellä tietynlaista lokiviestiä, lajitella sen attribuutit ja lisätä mm. värikoodattuna excel-tiedostoon.

Ohjelma ottaa syötteeksi loki-tiedoston, jonka sisältönä on tekstidataa, ja lajittelee sen exceliin taulukkomuotoon.
Tekstidatan koostuu lokiviestin alusta, 

Esim. 06:00:10.7301 Info {

Sekä Json-muotoisesta tekstidatasta, jonka rakenne on seuraava: 
"message":"some text ","level":"Information","logType":"Default","timeStamp": "....."

- Lajittelusta tekee haastavan se että lokien attirbuuttien sisällä olevat arvot voivat sisältää sisäkkäisiä attribuutteja; esim. Message kentän arvona saattaa lukea sisäkkäisenä toinen message-attribuutti, sekä sen arvo. Siinä tapauksessa lisätään se saman rivin message-soluun. Mikäli kyseessä on uusi attribuutti, lisätään se otsakkeisiin ja data oikean sarakkeen soluun. 

Ohjelma tarvitsee .NET-frameworkin asennettuna koneelle ennen .exe-tiedostoon kääntämistä

Ajo onnistuu komentoriviltä joko kahdella tai yhdellä parametrilla. 
Parametriksi sopii absoluuttinen tai suhteellinen polku ajokansiosta käsiteltävään tiedostoon.

![oneParam](https://user-images.githubusercontent.com/74860432/172547013-c6f821bf-1409-4879-901d-b9e4031cad97.png)

Yhden argumentin ajo takaa samannimisen tiedoston samaan kansioon, kuin missä argumentinkin tiedosto sijaitsee.
![twoParam](https://user-images.githubusercontent.com/74860432/172547017-4ed42224-029d-435b-b2e0-1c19a75f042d.png)

Toinen argumentti määrittää tuloste-tiedoston sijainnin, sekä nimen. Molemmat täytyy sisältää argumenttiin.
![invalidParam](https://user-images.githubusercontent.com/74860432/172547022-0ea5a975-2547-4213-9d5d-05fda1d4eb09.png)

Mikäli polku on virheellinen, tai tiedoston nimi väärä, palataan pääohjelmasta.
![success](https://user-images.githubusercontent.com/74860432/172547028-f229ba73-ac18-4c4c-903b-5a540c553528.png)

Pidemmät tiedostot, joissa dataa on useita satoja rivejä, sekä sarakkeita, hidastavat ajoa. 
Ajon päätyttyä oheinen teksti  muodostuu ruutun, jossa on kerrotuna excel-tiedoston sijainti.

![exampleProduct](https://user-images.githubusercontent.com/74860432/172547035-e4f47ad2-d4c8-4d12-9ab8-f2e0d3320e25.png)

Tuotetussa taulukossa on mm. Filtteröinti käytössä.
