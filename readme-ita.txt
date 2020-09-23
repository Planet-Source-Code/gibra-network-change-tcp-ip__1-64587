============================================================
 Network Change TCP/IP
 Giorgio Brausi - VBCorner
 http://www.vbcorner.net
============================================================
Questo progetto vuole essere un piccolo strumento di aiuto 
per gli sviluppatori che devono spesso modificare i parametri 
del protocollo TCP/IP del loro pc per connettersi, di volta 
in volta, ai server di rete di aziende diverse.
Ormai era diventata una paranoia dover introdurre i parametri 
a mano ogni volta che cambiavo azienda (e quindi rete).

Questo progetto non vuole essere né perfetto né completo, ma 
solo un punto di partenza che poi ognno può personalizzare in 
base alle proprie esigenze.

Il sorgente è quindi disponibile e modificabile senza alcuna 
limitazione.
Se modificate e/o correggete parte di questo progetto, vi sarò 
grato se vorrete segnalarmelo per l'opportuno aggiornamento sul 
sito, così che anche altri potranno usufruire dei miglioramenti 
apportati da voi, come voi, lo spero, state usufruendo di 
questo progetto.

Un ringraziamento a Mario Raccagni per il supporto tecnico.


============================================================
IMPORTANTE: 
Il programma di installazione crea un collegamento nella 
cartella "Esecuzione automatica", in questo modo Network 
Change sarà avviato automaticamente con Windows.
Dato che nel collegamento sarà specificato anche il parametro
HIDEONSTARTUP la sola cosa che vedrete sarà l'icona di Network 
Change nell'area di notifica della barra applicazioni.
Clic DESTRO sull'icona per aprire il menu di Network Change.
============================================================

VERSION HISTORY
============================================================
Versione 1.3.5 - 05/03/2006

        - NUOVO
          Aggiunto il file di lingua Tedesco. 
          Grazie a Patrik Menne!

        - NUOVO
          Aggiunta l'opzione "Auto-seleeiona la scheda di rete", 
          in questo modo se il vostro PC ha una sola scheda di
          rete la finestra 'Seleziona NIC' non è visualizzata
          e la vostra scheda è selezionata automaticamente. 
          
        - CHANGE - Parametro di comando /HIDEONSTARTUP:
          Il precedente parametro /AUTO è stato sostituito dal
          nuovo /HIDEONSTARTUP. Grazie a questo parametro, ora
          la finestra di NC non è più visualizzata all'avvio,
          ma sarà mostrata solo l'icona nell'area di notifica.

        - CHANGE - Finestra di NC
          La dimensione della finestra di NC window è stata
          allargata per consentire l'aggiunta di nuove lingua
          senza incorrere in problemi di dimensione delle 
          stringhe 

        - Attivando un profilo NC non chiede più conferma
	  per azzerare i precedenti parametri TCP/IP

	- i parametri di WINS primary e WINS secondary
	  sono ora tutti impostati a 0. 

	- Nuova finestra di About.

	
============================================================
Versione 1.3.2 - 08/02/2006
	- Quando si attiva un profilo sono prima azzerati
	  tutti i parametri, poi sono applicati quelli del
	  nuovo profilo. Ciò evita di mescolare parametri
	  vecchi e nuovi.

============================================================
Versione 1.3.1 - 18/01/2006
	- Se, quando si attiva un profilo, il computer ha una 
	  sola scheda di rete, tale scheda sarà selezionata
	  nella finestra "Seleziona schede di rete" in modo
	  automatico.

============================================================
Versione 1.3.0 - 08/10/2005
	Autore: Doretto Roberto
	- Aggiunto il form 'frmSelectNIC" per selezionare la
	  scheda di rete (qualora ne aveste più di una).
	- Aggiunte le stringhe nei files di linguaggio

	Ottimo lavoro Roberto. Thank!

============================================================
Versione 1.2.1 - 26/09/2005
	- alcuni bug corretti
	- alcuni piccoli miglioramenti

============================================================
Versione 1.1.0 - 21/06/2005
	- Il MaxLength dei txtIP è ora impostato su 3
	- Dopo aver digitato 3 numeri nei tetxbox txtIP 
	  il focus passa automaticamente al campo successivo.
	- Così avviene anche se si digita il punto '.'
	- Aggiunto il punto '.' tra i caratteri consentiti
	  nei textbox txtIP
	- Eliminato il controllo sull'esistenza del DNS
	  alternativo, che è infatti 'opzionale' e che
	  obbligava ad inserire almeno uno '0'.
	- spostato il controllo di alcuni tasti dall'evento
	  KeyDown all'evento KeyUp dei textbox txtIP

	Autore: Doretto Roberto
	Imposta automaticamente il Subnet Mask in base
	all'indirizzo IP
	
============================================================
Versione 1.0.0 - 13-06-2005

	Primo rilascio