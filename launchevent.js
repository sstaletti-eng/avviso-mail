/*
 * Avviso OPS - Spedimar
 * Avvisa quando tra i destinatari c'e' un indirizzo ESTERNO che assomiglia
 * a una casella interna Spedimar (es. "ops", "commerciali").
 *
 * - All'aggiunta di un destinatario (A, Cc, Ccn): barra rossa sopra il messaggio.
 * - All'invio: finestra con "Invia comunque" / "Non inviare".
 *
 * Per cambiare il comportamento modificare solo le due liste qui sotto.
 * Questo file gira sia in Outlook classico (runtime solo JavaScript) sia in
 * nuovo Outlook / Outlook sul web: niente import, niente sintassi moderna.
 */

// Domini considerati interni (anche i sottodomini, es. mail.spedimar.ch)
var DOMINI_INTERNI = ["spedimar.ch", "spedimarsagl.ch", "spedimar21.onmicrosoft.com"];

// Regole: "parola" e' cercata all'inizio di una parola nel nome visualizzato
// o nella parte prima della @. "interna" e' la casella Spedimar da citare nell'avviso.
// Per aggiungere una regola copiare una riga e cambiarla.
var REGOLE = [
  { parola: "ops",        interna: "operativi@spedimar.ch" },
  { parola: "commercial", interna: "commerciali@spedimar.ch" } // copre commerciali, commerciale, commercial
];

var ID_BARRA = "avvisoOpsEsterno";

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function iniziaParola(testo, parola) {
  var re = new RegExp("(^|[^a-z0-9])" + escapeRegex(parola), "i");
  return re.test(testo);
}

function isInterno(dominio) {
  for (var i = 0; i < DOMINI_INTERNI.length; i++) {
    var d = DOMINI_INTERNI[i];
    if (dominio === d || dominio.slice(-(d.length + 1)) === "." + d) {
      return true;
    }
  }
  return false;
}

// Ritorna la regola che scatta per questo destinatario, oppure null
function regolaSospetta(r) {
  var email = (r.emailAddress || "").toLowerCase().trim();
  var nome = (r.displayName || "").toLowerCase();
  var at = email.lastIndexOf("@");
  if (at < 1) {
    return null; // indirizzo non SMTP (es. voce interna Exchange): non e' esterno
  }
  var locale = email.substring(0, at);
  var dominio = email.substring(at + 1);
  if (isInterno(dominio)) {
    return null;
  }
  for (var i = 0; i < REGOLE.length; i++) {
    if (iniziaParola(locale, REGOLE[i].parola) || iniziaParola(nome, REGOLE[i].parola)) {
      return REGOLE[i];
    }
  }
  return null;
}

function leggiCampo(campo) {
  return new Promise(function (resolve) {
    if (!campo || !campo.getAsync) {
      resolve([]);
      return;
    }
    campo.getAsync(function (res) {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value || [] : []);
    });
  });
}

// Ritorna [{ email, interna }] senza doppioni
function trovaSospetti() {
  var item = Office.context.mailbox.item;
  return Promise.all([leggiCampo(item.to), leggiCampo(item.cc), leggiCampo(item.bcc)]).then(
    function (campi) {
      var tutti = campi[0].concat(campi[1], campi[2]);
      var visti = {};
      var sospetti = [];
      for (var i = 0; i < tutti.length; i++) {
        var email = (tutti[i].emailAddress || "").toLowerCase();
        var regola = regolaSospetta(tutti[i]);
        if (regola && !visti[email]) {
          visti[email] = true;
          sospetti.push({ email: email, interna: regola.interna });
        }
      }
      return sospetti;
    }
  );
}

function testoBarra(sospetti) {
  var testo;
  if (sospetti.length === 1) {
    testo = "ATTENZIONE: " + sospetti[0].email + " e' ESTERNO, non e' " + sospetti[0].interna;
  } else {
    var elenco = [];
    for (var i = 0; i < sospetti.length; i++) {
      elenco.push(sospetti[i].email);
    }
    testo = "ATTENZIONE: " + elenco.join(", ") + " sono ESTERNI, non caselle Spedimar";
  }
  if (testo.length > 150) {
    testo = "ATTENZIONE: " + sospetti.length + " indirizzi ESTERNI simili a caselle Spedimar";
  }
  return testo.substring(0, 150);
}

function aggiornaBarra(sospetti) {
  return new Promise(function (resolve) {
    var msgs = Office.context.mailbox.item.notificationMessages;
    if (sospetti.length > 0) {
      msgs.replaceAsync(
        ID_BARRA,
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: testoBarra(sospetti)
        },
        function () { resolve(); }
      );
    } else {
      msgs.removeAsync(ID_BARRA, function () { resolve(); });
    }
  });
}

// Evento: destinatari cambiati (A, Cc, Ccn)
function onRecipientsChanged(event) {
  trovaSospetti()
    .then(aggiornaBarra)
    .then(function () { event.completed(); })
    .catch(function () { event.completed(); });
}

// Evento: invio (Smart Alerts, SendMode = PromptUser)
function onMessageSend(event) {
  trovaSospetti()
    .then(function (sospetti) {
      if (sospetti.length === 0) {
        event.completed({ allowEvent: true });
        return;
      }
      var righe = [];
      for (var i = 0; i < sospetti.length; i++) {
        righe.push(sospetti[i].email + "  (NON e' " + sospetti[i].interna + ")");
      }
      event.completed({
        allowEvent: false,
        errorMessage:
          "Stai inviando a indirizzi ESTERNI che assomigliano a caselle Spedimar:\n\n" +
          righe.join("\n") +
          "\n\nSe e' davvero il cliente/fornitore giusto, scegli 'Invia comunque'."
      });
    })
    .catch(function () {
      // In caso di errore non blocchiamo mai l'invio
      event.completed({ allowEvent: true });
    });
}

// Pulsante "Controlla destinatari" nella barra del messaggio (controllo manuale / verifica installazione)
function controllaOra(event) {
  trovaSospetti()
    .then(function (sospetti) {
      if (sospetti.length > 0) {
        return aggiornaBarra(sospetti);
      }
      return new Promise(function (resolve) {
        Office.context.mailbox.item.notificationMessages.replaceAsync(
          ID_BARRA,
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Avviso OPS attivo: nessun destinatario esterno sospetto.",
            icon: "Icon.16x16",
            persistent: false
          },
          function () { resolve(); }
        );
      });
    })
    .then(function () { event.completed(); })
    .catch(function () { event.completed(); });
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onRecipientsChanged", onRecipientsChanged);
  Office.actions.associate("onMessageSend", onMessageSend);
  Office.actions.associate("controllaOra", controllaOra);
}
