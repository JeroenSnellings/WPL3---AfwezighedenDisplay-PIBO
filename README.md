Afwezigheidsmonitoringssysteem

## Overzicht

Dit project is verantwoordelijk voor het monitoren en weergeven van afwezigheden van docenten van de PIBO in de inkomhal van de school. Wanneer een medewerker van het secretariaat afwezigheden van docenten invult in het gedeelde Excel-bestand "afwezigheden.xlsx," zal de Raspberry Pi-module, die dit script uitvoert, wijzigingen opmerken en deze automatisch bijwerken op het monitoringsdisplay in de schoolgang.

## Projectproces

Hier is een overzicht van hoe dit project werkt:

1. **Excel-bestand bijwerken**: Een lid van de administratie vult afwezigheidsinformatie in voor docenten in het gedeelde Excel-bestand "afwezigheden.xlsx." Dit bestand wordt op een gedeelde locatie opgeslagen op de server, zodat het toegankelijk is voor zowel de administratie als de Raspberry Pi-module.

2. **Raspberry Pi-module**: De Raspberry Pi is een kleine, krachtige computer die dit project aandrijft. Op de Raspberry Pi wordt een Python-script uitgevoerd dat de gedeelde Excel-gegevens controleert.

3. **Python-script**: Het Python-script in dit project maakt gebruik van de Tkinter-bibliotheek om een venster te maken dat de afwezigheidsinformatie weergeeft. Het script controleert voortdurend of het Excel-bestand is gewijzigd.

4. **Bestandscontrole**: Het script maakt gebruik van het besturingssysteem om de laatst gewijzigde tijd van het Excel-bestand "afwezigheden.xlsx" te controleren. Als deze tijd verandert, betekent dit dat er nieuwe gegevens zijn ingevoerd.

5. **Automatische updates**: Als er wijzigingen worden gedetecteerd, werkt het script het venster bij met de nieuwe gegevens uit het Excel-bestand. Het venster wordt volledig vernieuwd met de actuele afwezigheidsinformatie.

6. **Fullscreen**: Het script draait in Fullscreen om de informatie duidelijk en leesbaar weer te geven op het monitoringsdisplay in de schoolgang.

7. **Uit de volledig schermmodus gaan**: Voor debuggen is er een mogelijkheid om uit de volledig schermmodus te gaan door op de 'Escape'-toets te drukken op een toetsenbord dat aangesloten wordt op de raspberry Pi.

## Auteursrecht en Toestemming

Dit project is gemaakt door Jeroen Snellings en is auteursrechtelijk beschermd. Hierbij wordt toestemming verleend om dit project kosteloos te gebruiken, te kopiëren, te wijzigen, en te verspreiden onder de voorwaarden zoals beschreven in de bovenstaande toestemmingsverklaring.

## Opmerking

Het verkopen van dit project of afgeleide werken is niet toegestaan. Het project wordt geleverd "as is," zonder enige garantie. De auteur is niet aansprakelijk voor enige claim, schade of andere aansprakelijkheid voortvloeiend uit het gebruik van de software.

---

Voor vragen of ondersteuning, neem contact op met Jeroen Snellings.



*****************************
*							*
*  Raspberry pi Credentials	*
*							*
*		Raspberrypi.local	*
*		user: pi			*	
*		pass: raspberry		*
*****************************