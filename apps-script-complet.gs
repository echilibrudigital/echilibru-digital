// ═══════════════════════════════════════════════════════════════════════
// ECHILIBRU DIGITAL - APPS SCRIPT COMPLET (REORGANIZAT)
// ═══════════════════════════════════════════════════════════════════════
// Versiune: Reorganizată - 16.02.2026
// Structură:
//   1. CONFIGURARE
//   2. STRIPE API
//   3. FUNCȚII HELPER GENERALE
//   4. SINCRONIZARE AUTOMATĂ PLĂȚI
//   5. VERIFICARE DUPLICATE
//   6. MANAGEMENT CLIENȚI
//   7. PROCESARE CURSURI (Curs 1, Curs 2, Pachet)
//   8. SISTEM REFERRAL
//   9. FACTURARE PDF
//  10. EMAILURI ACCES CURSURI
//  11. API WEB (doGet)
//  12. PROCESARE INSTANTĂ STRIPE
//  13. SISTEM SECURITATE (Token + Cod Verificare)
//  14. SISTEM TRIAL 7 ZILE
//  15. ÎNREGISTRARE TRIAL MANUAL
//  16. AUTOMATIZARE EMAILURI TRIAL (Ziua 4, Ziua 7)
//  17. SISTEM PLĂȚI COMISIOANE
//  18. DEBUG & TESTARE
// ═══════════════════════════════════════════════════════════════════════


// ═════════════════════════════════════════════════════════
// 1. CONFIGURARE
// ═════════════════════════════════════════════════════════

var TEMPLATE_CURS1_ID = '1RjYNW7oGlY15yMQbtUlCaHzM7K1r2euAGSuPblr5YIE';
var TEMPLATE_CURS2_ID = '1eHN-1zkXoRkmM4S9WiYGK60rYsJ79rr4NqFtNg_RJPk';
var TEMPLATE_PACHET_ID = '1aVUOlHvovOG1e518c1zozJxNvw-1_jL_QFFXzc8Udws';
var TEMPLATE_PV_PF_ID = '1qD6Fr3wbPmJHTze_pP9DGXcmzao1jN654w5SlmqsDoI'; // Template PV Persoana Fizica
var TEMPLATE_PV_PJ_ID = '1rGPc8JSlE770-2oQawB2owP0jOeoUAHohFmk-hPC3os'; // Template PV Firma
var FOLDER_PV_ARHIVA_ID = '1ZllgtH-SzP2kIORHAMdbyDeGFFHrj9Bi'; // Folder Arhiva PV
var SITE_URL = 'https://echilibrudigital.ro';


// ═════════════════════════════════════════════════════════
// 2. STRIPE API
// ═════════════════════════════════════════════════════════

function setareStripeKey() {
  var key = 'sk_live_(codul meu)';
  PropertiesService.getScriptProperties().setProperty('STRIPE_SECRET_KEY', key);
  Logger.log('✅ Stripe Key salvat cu succes!');
}

function getStripeKey() {
  var key = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!key) {
    throw new Error('⚠️ Lipsește STRIPE_SECRET_KEY! Rulează setareStripeKey() mai întâi.');
  }
  return key;
}

function stripeApiRequest(endpoint, method, params) {
  var apiKey = getStripeKey();
  var url = 'https://api.stripe.com/v1/' + endpoint;
  
  var options = {
    method: method || 'get',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    muteHttpExceptions: true
  };
  
  if (params && method === 'post') {
    options.payload = params;
  }
  
  if (params && method === 'get') {
    var queryParts = [];
    for (var key in params) {
      var value = params[key];
      if (typeof value === 'object' && value !== null) {
        for (var subKey in value) {
          queryParts.push(key + '[' + subKey + ']=' + encodeURIComponent(value[subKey]));
        }
      } else {
        queryParts.push(key + '=' + encodeURIComponent(value));
      }
    }
    if (queryParts.length > 0) {
      url += '?' + queryParts.join('&');
    }
  }
  
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  
  if (result.error) {
    throw new Error('Stripe API Error: ' + result.error.message);
  }
  
  return result;
}


// ═════════════════════════════════════════════════════════
// 3. FUNCȚII HELPER GENERALE
// ═════════════════════════════════════════════════════════

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}

function genereazaParola() {
  var caractere = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var parola = '';
  for (var i = 0; i < 10; i++) {
    parola += caractere.charAt(Math.floor(Math.random() * caractere.length));
  }
  return parola;
}

function generareNumarFactura() {
  var sheet = getOrCreateSheet('Contor Facturi', ['Ultimul Numar']);
  var numarCurent = sheet.getRange('A2').getValue() || 0;
  var numarNou = parseInt(numarCurent) + 1;
  sheet.getRange('A2').setValue(numarNou);
  return 'DDA-' + ('00000' + numarNou).slice(-5);
}

function generareNumarPV() {
  var sheet = getOrCreateSheet('Contor PV', ['Ultimul Numar PV']);
  var numarCurent = sheet.getRange('A2').getValue() || 0;
  var numarNou = parseInt(numarCurent) + 1;
  sheet.getRange('A2').setValue(numarNou);
  return 'PV-' + ('00000' + numarNou).slice(-5);
}

function extragePrenum(email) {
  try {
    var parte = email.split('@')[0];
    var nume = parte.split('.')[0];
    nume = nume.replace(/[0-9_-]/g, '');
    if (nume.length <= 1) {
      nume = parte.replace(/[0-9_-]/g, '');
    }
    if (nume.length > 0) {
      return nume.charAt(0).toUpperCase() + nume.slice(1).toLowerCase();
    }
    return 'Client';
  } catch(error) {
    return 'Client';
  }
}



// ═════════════════════════════════════════════════════════
// 4. SINCRONIZARE AUTOMATĂ PLĂȚI STRIPE
// ═════════════════════════════════════════════════════════

function sincronizarePlatiStripe() {
  var startTime = new Date().getTime();
  var MAX_DURATION = 300000;

  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var lastCheckProp = scriptProps.getProperty('LAST_STRIPE_CHECK');
    var lookupSince;

    if (lastCheckProp && !isNaN(parseInt(lastCheckProp))) {
      var timeSinceLastCheck = Math.floor(Date.now() / 1000) - parseInt(lastCheckProp);
      if (timeSinceLastCheck < 600) {
        Logger.log('Verificare prea recentă (' + timeSinceLastCheck + 's ago), skip.');
        return;
      }
    }

    if (!lastCheckProp || lastCheckProp === '' || isNaN(parseInt(lastCheckProp))) {
      lookupSince = Math.floor(Date.now() / 1000) - (3 * 24 * 3600);
      Logger.log('Prima rulare / timestamp invalid - caut ultimele 3 zile');
    } else if (Date.now() / 1000 - parseInt(lastCheckProp) > 30 * 24 * 3600) {
      lookupSince = Math.floor(Date.now() / 1000) - (7 * 24 * 3600);
      Logger.log('Timestamp vechi (>30 zile) - caut ultimele 7 zile');
    } else {
      lookupSince = parseInt(lastCheckProp);
      var maxLookback = Math.floor(Date.now() / 1000) - (14 * 24 * 3600);
      if (lookupSince < maxLookback) {
        lookupSince = maxLookback;
        Logger.log('LastCheck prea vechi - limitat la ultimele 14 zile');
      } else {
        Logger.log('Caz normal - caut de la: ' + new Date(lookupSince * 1000).toLocaleString('ro-RO'));
      }
    }

    var allPaidSessions = [];
    var startingAfter = null;
    var hasMore = true;
    var safetyCounter = 0;

    while (hasMore && (new Date().getTime() - startTime) < MAX_DURATION && safetyCounter < 100) {
      safetyCounter++;
      var params = { 
        limit: 100, 
        created: { gte: lookupSince },
        status: 'complete'
      };
      if (startingAfter) params.starting_after = startingAfter;

      var page = stripeApiRequest('checkout/sessions', 'get', params);

      if (page.data.length > 0) {
        allPaidSessions = allPaidSessions.concat(page.data);
      }

      hasMore = page.has_more && page.data.length === 100;
      if (hasMore) startingAfter = page.data[page.data.length - 1].id;
    }

    if (allPaidSessions.length === 0) {
      Logger.log('Nu sunt plăți noi de procesat.');
      scriptProps.setProperty('LAST_STRIPE_CHECK', Math.floor(Date.now() / 1000).toString());
      return;
    }

    var procesate = 0;
    var duplicate = 0;

    for (var i = 0; i < allPaidSessions.length; i++) {
      if ((new Date().getTime() - startTime) > MAX_DURATION) {
        Logger.log('Aproape de timeout - opresc procesarea.');
        break;
      }

      var session = allPaidSessions[i];
      var sessionId = session.id;

      var sessionLock = LockService.getScriptLock();
      try {
        sessionLock.waitLock(10000);
      } catch (e) {
        Logger.log('Sincronizarea a sarit peste ' + sessionId + ' din cauza lock-ului ocupat.');
        continue;
      }

      try {
        if (verificaDuplicateSession(sessionId)) {
          duplicate++;
          sessionLock.releaseLock();
          continue;
        }

        var email = session.customer_details?.email || 'necunoscut@email.com';
        var nume  = session.customer_details?.name  || 'Client';
        var tipProdus = (session.metadata || {}).tip_produs;
        var pret = session.amount_total / 100;
        var valuta = (session.currency || 'ron').toUpperCase();

        Logger.log('PROCESEZ: ' + email + ' | ' + tipProdus + ' | ' + pret + ' ' + valuta + ' | ' + sessionId);
        marcheazaSessionProcesat(sessionId);
        sessionLock.releaseLock();

        try {
          if (tipProdus === 'echilibru_digital')          procesareCurs1(email, nume, pret, valuta, sessionId, session.metadata);
          else if (tipProdus === 'curs_parinti')          procesareCurs2(email, nume, pret, valuta, sessionId, session.metadata);
          else if (tipProdus === 'pachet_redus')          procesarePachetComplet(email, nume, pret, valuta, sessionId, session.metadata);
          else Logger.log('Tip produs necunoscut: ' + tipProdus);
          procesate++;
        } catch (procErr) {
          Logger.log('Eroare procesare individuala ' + sessionId + ': ' + procErr.toString());
        }

      } catch (err) {
        if (sessionLock.hasLock()) sessionLock.releaseLock();
        Logger.log('Eroare generala sesiune ' + sessionId + ': ' + err.toString());
      }
    }

    scriptProps.setProperty('LAST_STRIPE_CHECK', Math.floor(Date.now() / 1000).toString());
    var durata = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    Logger.log('SINCRONIZARE REUSITA! Procesate: ' + procesate + ', Duplicate: ' + duplicate + ', Durata: ' + durata + 's');

  } catch (err) {
    Logger.log('EROARE CRITICA: ' + err.toString());
    
    var scriptProps = PropertiesService.getScriptProperties();
    var lastErrorEmail = scriptProps.getProperty('LAST_ERROR_EMAIL');
    var now = Date.now();
    
    if (!lastErrorEmail || (now - parseInt(lastErrorEmail)) > 3600000) {
      MailApp.sendEmail({
        to: 'gandulsanatatii@gmail.com',
        subject: 'CRITICAL: Script Stripe a cazut!',
        body: 'Data: ' + new Date() + '\nEroare: ' + err.toString() + '\n\nStack:\n' + err.stack
      });
      scriptProps.setProperty('LAST_ERROR_EMAIL', now.toString());
    }
  }
}

function creeazaTriggerSincronizare() {
  var triggere = ScriptApp.getProjectTriggers();
  triggere.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sincronizarePlatiStripe') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('sincronizarePlatiStripe')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  Logger.log('Trigger creat! Sincronizare automata la fiecare 15 minute');
}


// ═════════════════════════════════════════════════════════
// 5. VERIFICARE DUPLICATE SESSION
// ═════════════════════════════════════════════════════════

function verificaDuplicateSession(sessionId) {
  var sheet = getOrCreateSheet('Sessions Procesate', ['Session ID', 'Data', 'Status']);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId) return true;
  }
  return false;
}

function marcheazaSessionProcesat(sessionId) {
  var sheet = getOrCreateSheet('Sessions Procesate', ['Session ID', 'Data', 'Status']);
  sheet.appendRow([sessionId, new Date(), 'Procesat']);
}


// ═════════════════════════════════════════════════════════
// 6. MANAGEMENT CLIENTI
// ═════════════════════════════════════════════════════════

function cautaClientExistent(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet || sheet.getLastRow() < 2) return null;
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      return {
        rand: i + 1,
        email: data[i][0],
        nume: data[i][1],
        parola: data[i][3],
        accesCurs1: data[i][4],
        accesCurs2: data[i][5],
        accesPachet: data[i][6]
      };
    }
  }
  
  return null;
}

function actualizareAccesCurs(email, tipCurs, valoare) {
  var client = cautaClientExistent(email);
  if (!client) return;
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  
  if (tipCurs === 'curs1') {
    sheet.getRange(client.rand, 5).setValue(valoare);
  } else if (tipCurs === 'curs2') {
    sheet.getRange(client.rand, 6).setValue(valoare);
  } else if (tipCurs === 'pachet') {
    sheet.getRange(client.rand, 5).setValue(true);
    sheet.getRange(client.rand, 6).setValue(true);
    sheet.getRange(client.rand, 7).setValue(true);
  }
  
  sheet.getRange(client.rand, 8).setValue(new Date());
  Logger.log('Acces actualizat pentru ' + email);
}



// ═════════════════════════════════════════════════════════
// 7. PROCESARE CURS 1 (Echilibru Digital)
// ═════════════════════════════════════════════════════════

function procesareCurs1(email, nume, pret, valuta, sessionId, metadata) {
  try {
    Logger.log('[CURS1] Procesare: ' + email);
    
    var infoTrial = detectareCuponTrial(sessionId);
    var tipAcces = infoTrial.tipAcces;
    var dataExpirare = infoTrial.dataExpirare;
    
    Logger.log('[CURS1] Tip acces: ' + tipAcces);
    if (infoTrial.esteTrial) {
      Logger.log('[CURS1] TRIAL DETECTAT! Expirare: ' + dataExpirare);
    }
    
    var clientExistent = cautaClientExistent(email);
    var parola;
    
    if (clientExistent) {
      Logger.log('[CURS1] Client existent gasit.');
      parola = clientExistent.parola;
      
      actualizareAccesCurs(email, 'curs1', true);
      
      var clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
      if (clientSheet) {
        var data = clientSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] === email) {
            clientSheet.getRange(i + 1, 10).setValue(tipAcces);
            if (dataExpirare) {
              clientSheet.getRange(i + 1, 11).setValue(dataExpirare);
            }
            break;
          }
        }
      }
      
      if (tipAcces === 'Platit') {
        actualizeazaTrialLaConvertit(email, 'curs1');
      }
      
    } else {
      Logger.log('[CURS1] Client nou - generez cont.');
      parola = genereazaParola();
      
      var sheetClienti = getOrCreateSheet('Clienti', [
        'Email', 'Nume', 'Data Inscriere', 'Parola', 
        'Acces Curs 1', 'Acces Curs 2', 'Acces Pachet',
        'Data ultima achizitie', 'Total Achizitii',
        'Tip Acces', 'Data Expirare Trial'
      ]);
      
      sheetClienti.appendRow([
        email, nume, new Date(), parola,
        true, false, false,
        new Date(),
        '',
        tipAcces,
        dataExpirare || ''
      ]);
    }
    
    if (infoTrial.esteTrial) {
      salveazaTrial(email, 'curs1', infoTrial.cuponId, new Date(), dataExpirare);
    }
    
    var numarFactura = generareNumarFactura();
    
    var sheetFacturi = getOrCreateSheet('Facturi Curs', [
      'Nr Factura', 'Data', 'Client', 'Email', 'Tip Curs', 'Pret', 'Status', 'Session ID'
    ]);
    sheetFacturi.appendRow([
      numarFactura, new Date(), nume, email, 'Curs 1', pret + ' ' + valuta, 'Trimisa', sessionId
    ]);
    
    // Tracking Referral
    var codReferral = null;
    if (metadata && metadata.referral_code) {
      codReferral = metadata.referral_code;
    }
    
    if (!codReferral) {
      try {
        var sessionComplet = stripeApiRequest('checkout/sessions/' + sessionId, 'get', {});
        if (sessionComplet.client_reference_id) {
          codReferral = sessionComplet.client_reference_id;
        }
      } catch(e) {}
    }
    
    if (codReferral && codReferral !== 'null' && codReferral !== '') {
      inregistreazaConversieReferral(codReferral, email, pret, valuta);
    }
    
    trimiteEmailAccesCurs1(email, nume, parola, clientExistent, tipAcces, dataExpirare);
    genereazaSiTrimiteFactura(email, nume, pret, valuta, numarFactura, 'curs1');
    
    Logger.log('[CURS1] Finalizat cu succes! Tip: ' + tipAcces);
    return { success: true, parola: parola, email: email, tipAcces: tipAcces };
    
  } catch(error) {
    Logger.log('[EROARE CURS1] ' + error.toString());
    throw error;
  }
}


// ═════════════════════════════════════════════════════════
// 7b. PROCESARE CURS 2 (Curs Parinti)
// ═════════════════════════════════════════════════════════

function procesareCurs2(email, nume, pret, valuta, sessionId, metadata) {
  try {
    Logger.log('[CURS2] Procesare: ' + email);
    
    var infoTrial = detectareCuponTrial(sessionId);
    var tipAcces = infoTrial.tipAcces;
    var dataExpirare = infoTrial.dataExpirare;
    
    if (infoTrial.esteTrial) {
      Logger.log('[CURS2] TRIAL DETECTAT! Expirare: ' + dataExpirare);
    }
    
    var clientExistent = cautaClientExistent(email);
    var parola;
    
    if (clientExistent) {
      parola = clientExistent.parola;
      actualizareAccesCurs(email, 'curs2', true);
      
      var clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
      if (clientSheet) {
        var data = clientSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] === email) {
            clientSheet.getRange(i + 1, 10).setValue(tipAcces);
            if (dataExpirare) {
              clientSheet.getRange(i + 1, 11).setValue(dataExpirare);
            }
            break;
          }
        }
      }
      
      if (tipAcces === 'Platit') {
        actualizeazaTrialLaConvertit(email, 'curs_parinti');
      }
      
    } else {
      parola = genereazaParola();
      
      var sheetClienti = getOrCreateSheet('Clienti', [
        'Email', 'Nume', 'Data Inscriere', 'Parola', 
        'Acces Curs 1', 'Acces Curs 2', 'Acces Pachet',
        'Data ultima achizitie', 'Total Achizitii',
        'Tip Acces', 'Data Expirare Trial'
      ]);
      
      sheetClienti.appendRow([
        email, nume, new Date(), parola,
        false, true, false,
        new Date(),
        '',
        tipAcces,
        dataExpirare || ''
      ]);
    }
    
    if (infoTrial.esteTrial) {
      salveazaTrial(email, 'curs_parinti', infoTrial.cuponId, new Date(), dataExpirare);
    }
    
    var numarFactura = generareNumarFactura();
    
    var sheetFacturi = getOrCreateSheet('Facturi Curs', [
      'Nr Factura', 'Data', 'Client', 'Email', 'Tip Curs', 'Pret', 'Status', 'Session ID'
    ]);
    sheetFacturi.appendRow([
      numarFactura, new Date(), nume, email, 'Curs 2', pret + ' ' + valuta, 'Trimisa', sessionId
    ]);
    
    var codReferral = null;
    if (metadata && metadata.referral_code) codReferral = metadata.referral_code;
    if (!codReferral) {
      try {
        var sessionComplet = stripeApiRequest('checkout/sessions/' + sessionId, 'get', {});
        if (sessionComplet.client_reference_id) codReferral = sessionComplet.client_reference_id;
      } catch(e) {}
    }
    if (codReferral && codReferral !== 'null' && codReferral !== '') {
      inregistreazaConversieReferral(codReferral, email, pret, valuta);
    }
    
    trimiteEmailAccesCurs2(email, nume, parola, clientExistent, tipAcces, dataExpirare);
    genereazaSiTrimiteFactura(email, nume, pret, valuta, numarFactura, 'curs2');
    
    Logger.log('[CURS2] Finalizat cu succes! Tip: ' + tipAcces);
    return { success: true, parola: parola, email: email, tipAcces: tipAcces };
    
  } catch(error) {
    Logger.log('[EROARE CURS2] ' + error.toString());
    throw error;
  }
}


// ═════════════════════════════════════════════════════════
// 7c. PROCESARE PACHET COMPLET (Curs 1 + Curs 2)
// ═════════════════════════════════════════════════════════

function procesarePachetComplet(email, nume, pret, valuta, sessionId, metadata) {
  try {
    Logger.log('[PACHET] Procesare: ' + email);
    
    var clientExistent = cautaClientExistent(email);
    var parola;
    
    if (clientExistent) {
      parola = clientExistent.parola;
      actualizareAccesCurs(email, 'pachet', true);
      
      var clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
      if (clientSheet) {
        var data = clientSheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] && data[i][0].toString().toLowerCase() === email.toLowerCase()) {
            clientSheet.getRange(i + 1, 10).setValue('Platit');
            break;
          }
        }
      }
    } else {
      parola = genereazaParola();
      
      var sheetClienti = getOrCreateSheet('Clienti', [
        'Email', 'Nume', 'Data Inscriere', 'Parola', 
        'Acces Curs 1', 'Acces Curs 2', 'Acces Pachet',
        'Data ultima achizitie', 'Total Achizitii',
        'Tip Acces', 'Data Expirare Trial'
      ]);
      
      sheetClienti.appendRow([
        email, nume, new Date(), parola,
        true, true, true,
        new Date(),
        '',
        'Platit',
        ''
      ]);
    }
    
    var numarFactura = generareNumarFactura();
    
    var sheetFacturi = getOrCreateSheet('Facturi Curs', [
      'Nr Factura', 'Data', 'Client', 'Email', 'Tip Curs', 'Pret', 'Status', 'Session ID'
    ]);
    sheetFacturi.appendRow([
      numarFactura, new Date(), nume, email, 'Pachet complet', pret + ' ' + valuta, 'Trimisa', sessionId
    ]);
    
    var codReferral = null;
    if (metadata && metadata.referral_code) codReferral = metadata.referral_code;
    if (!codReferral) {
      try {
        var sessionComplet = stripeApiRequest('checkout/sessions/' + sessionId, 'get', {});
        if (sessionComplet.client_reference_id) codReferral = sessionComplet.client_reference_id;
      } catch(e) {}
    }
    if (codReferral && codReferral !== 'null' && codReferral !== '') {
      inregistreazaConversieReferral(codReferral, email, pret, valuta);
    }
    
    trimiteEmailAccesPachet(email, nume, parola, clientExistent);
    genereazaSiTrimiteFactura(email, nume, pret, valuta, numarFactura, 'pachet');
    
    Logger.log('[PACHET] Finalizat cu succes!');
    return { success: true, parola: parola, email: email, tipAcces: 'Platit' };
    
  } catch(error) {
    Logger.log('[EROARE PACHET] ' + error.toString());
    throw error;
  }
}


// ═════════════════════════════════════════════════════════
// 8. SISTEM REFERRAL
// ═════════════════════════════════════════════════════════

function genereazaLinkReferral(email) {
  try {
    var sheet = getOrCreateSheet('Referrals', [
      'Email Client', 'Cod Referral', 'Data Generare', 
      'Nr Conversii', 'Clienti Recomandati', 'Total Castigat (RON)',
      'Metoda Plata', 'Detalii Plata', 'RON Primiti'
    ]);
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        var codExistent = data[i][1];
        var stripeLinkCuReferral = 'https://buy.stripe.com/6oU00c00X8DT7qn6t2co000?client_reference_id=' + codExistent;
        
        // -------------------------------------------------------------------------
        // LOGICA LINK DINAMIC: Determinam URL-ul in functie de accesul clientului
        // -------------------------------------------------------------------------
        var baseUrl = SITE_URL; // Default: https://echilibrudigital.ro
        
        // Cautam clientul in baza de date Clienti pentru a vedea ce produse are
        var sheetClienti = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
        if (sheetClienti) {
          var dataClienti = sheetClienti.getDataRange().getValues();
          for (var k = 1; k < dataClienti.length; k++) {
            if (dataClienti[k][0] === email) {
              var accesCurs1 = dataClienti[k][4] === true || dataClienti[k][4] === 'TRUE';
              var accesCurs2 = dataClienti[k][5] === true || dataClienti[k][5] === 'TRUE';
              var accesPachet = dataClienti[k][6] === true || dataClienti[k][6] === 'TRUE';
              
              if (accesPachet || (accesCurs1 && accesCurs2)) {
                // Are totul -> Link general (homepage) unde userul alege
                baseUrl = SITE_URL;
              } else if (accesCurs1) {
                // Are doar Curs 1 -> Link specific Curs 1
                baseUrl = SITE_URL + '/mai-prezenta-nu-doar-conectata';
              } else if (accesCurs2) {
                // Are doar Curs 2 -> Link specific Curs 2
                baseUrl = SITE_URL + '/ghidul-parintelui-in-era-digitala';
              }
              break;
            }
          }
        }
        
        return {
          success: true,
          referralCode: codExistent,
          referralLink: baseUrl + '?ref=' + codExistent,
          stripeLinkDirect: stripeLinkCuReferral,
          conversii: data[i][3] || 0,
          totalCastigat: data[i][5] || 0,
          metodaPlata: data[i][6] || null,
          detaliiPlata: data[i][7] || null,
          ronPrimiti: data[i][8] || 0,
          cnp: data[i][9] || '',
          adresa: data[i][10] || '',
          firma: data[i][11] || '',
          cui: data[i][12] || ''
        };
      }
    }
    
    // -------------------------------------------------------------------------
    // LOGICA LINK DINAMIC: Determinam URL-ul in functie de accesul clientului
    // -------------------------------------------------------------------------
    var baseUrl = SITE_URL; // Default: https://echilibrudigital.ro
    
    // Cautam clientul in baza de date Clienti pentru a vedea ce produse are
    var sheetClienti = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
    if (sheetClienti) {
      var dataClienti = sheetClienti.getDataRange().getValues();
      for (var k = 1; k < dataClienti.length; k++) {
        if (dataClienti[k][0] === email) {
          var accesCurs1 = dataClienti[k][4] === true || dataClienti[k][4] === 'TRUE';
          var accesCurs2 = dataClienti[k][5] === true || dataClienti[k][5] === 'TRUE';
          var accesPachet = dataClienti[k][6] === true || dataClienti[k][6] === 'TRUE';
          
          if (accesPachet || (accesCurs1 && accesCurs2)) {
            // Are totul -> Link general (homepage) unde userul alege
            baseUrl = SITE_URL; 
          } else if (accesCurs1) {
            // Are doar Curs 1 -> Link specific Curs 1
            baseUrl = SITE_URL + '/mai-prezenta-nu-doar-conectata';
          } else if (accesCurs2) {
            // Are doar Curs 2 -> Link specific Curs 2
            baseUrl = SITE_URL + '/ghidul-parintelui-in-era-digitala';
          }
          break;
        }
      }
    }
    
    var codNou = genereazaCodReferralUnic();
    sheet.appendRow([email, codNou, new Date(), 0, '', 0, '', '', 0]);
    
    var stripeLinkCuReferral = 'https://buy.stripe.com/6oU00c00X8DT7qn6t2co000?client_reference_id=' + codNou;
    
    return {
      success: true,
      referralCode: codNou,
      referralLink: baseUrl + '?ref=' + codNou,
      stripeLinkDirect: stripeLinkCuReferral,
      conversii: 0,
      totalCastigat: 0,
      metodaPlata: null,
      detaliiPlata: null,
      ronPrimiti: 0,
      cnp: '',
      adresa: '',
      firma: '',
      cui: ''
    };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

function genereazaCodReferralUnic() {
  var caractere = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  var cod = 'REF';
  for (var i = 0; i < 8; i++) {
    cod += caractere.charAt(Math.floor(Math.random() * caractere.length));
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Referrals');
  if (sheet && sheet.getLastRow() > 1) {
    var coduri = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
    if (coduri.indexOf(cod) !== -1) {
      return genereazaCodReferralUnic();
    }
  }
  
  return cod;
}

function inregistreazaConversieReferral(codReferral, emailClientNou, pretAchizitie, valuta) {
  try {
    if (!codReferral) return;
    
    var sheetReferrals = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Referrals');
    if (!sheetReferrals) return;
    
    var data = sheetReferrals.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === codReferral) {
        var emailReferrer = data[i][0];
        var conversiiCurente = data[i][3] || 0;
        var clientiRecomandati = data[i][4] || '';
        var totalCastigat = data[i][5] || 0;
        
        var comision = 50; 
        var mesajComision = '50 RON'; 
        var status = 'Confirmat';
        
        // COMISION NELIMITAT - am scos limita de 12 conversii (600 RON)
        // if (conversiiCurente >= 12) { ... }
        
        sheetReferrals.getRange(i + 1, 4).setValue(conversiiCurente + 1);
        
        var listaNoua = clientiRecomandati 
          ? clientiRecomandati + ', ' + emailClientNou 
          : emailClientNou;
        sheetReferrals.getRange(i + 1, 5).setValue(listaNoua);
        sheetReferrals.getRange(i + 1, 6).setValue(totalCastigat + comision);
        
        var sheetConversii = getOrCreateSheet('Conversii Referral', [
          'Data', 'Email Referrer', 'Cod Referral', 'Email Client Nou',
          'Pret Achizitie', 'Valuta', 'Comision (50 RON)', 'Status'
        ]);
        
        sheetConversii.appendRow([
          new Date(), emailReferrer, codReferral, emailClientNou,
          pretAchizitie, valuta, mesajComision, status
        ]);
        
        if (comision > 0) {
          trimiteNotificareConversie(emailReferrer, emailClientNou, comision, valuta);
        }
        
        break;
      }
    }
    
  } catch(error) {
    Logger.log('[EROARE REFERRAL] ' + error.toString());
  }
}

function trimiteNotificareConversie(emailReferrer, emailClientNou, comision, valuta) {
  try {
    MailApp.sendEmail({
      to: emailReferrer,
      subject: 'Felicitări! Recomandarea ta a dat roade! 🎉',
      body: 'Bună,\n\nO veste excelentă: Cineva a achiziționat cursul folosind link-ul tău de recomandare!\n\nComision câștigat: 50 RON\n\nTe rugăm să accesezi dashboard-ul tău de recomandări și să salvezi metoda prin care dorești să primești plata (dacă nu ai făcut-o deja): https://echilibrudigital.ro/recomanda-cursul\n\nDupă salvarea metodei de plată, vei primi comisionul conform regulamentului.\n\nPentru detalii suplimentare despre distribuirea comisioanelor, te invităm să consulți regulamentul: https://docs.google.com/document/d/1Lk6-UnLYE6Pae50T3y8DNs0-vE3PLzD0VsXyYURn4Do/edit?usp=sharing\n\nÎți mulțumim pentru încredere și recomandare!\n\nCu respect,\nEchipa Echilibru Digital',
      name: 'Echilibru Digital'
    });
  } catch(error) {
    Logger.log('[EROARE] Trimitere notificare: ' + error.toString());
  }
}

function getStatisticiReferral(email) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Referrals');
    if (!sheet) {
      return { success: false, error: 'Nu exista date referral' };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        return {
          success: true,
          conversii: data[i][3] || 0,
          totalCastigat: data[i][5] || 0,
          clientiRecomandati: data[i][4] || '',
          codReferral: data[i][1]
        };
      }
    }
    
    return { success: false, error: 'Nu exista date pentru acest email' };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

function salveazaMetodaPlata(email, metoda, detalii, cnp, adresa, firma, cui) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Referrals');
    if (!sheet) {
      return { success: false, error: 'Nu exista tabel Referrals' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.getRange(i + 1, 7).setValue(metoda);
        sheet.getRange(i + 1, 8).setValue(detalii);
        
        // Campuri noi (col 10-14)
        sheet.getRange(i + 1, 10).setValue(cnp || '');
        sheet.getRange(i + 1, 11).setValue(adresa || '');
        sheet.getRange(i + 1, 12).setValue(firma || '');
        sheet.getRange(i + 1, 13).setValue(cui || '');
        
        var tipPlatitor = (cui && cui.length > 2) ? 'PJ' : 'PF';
        sheet.getRange(i + 1, 14).setValue(tipPlatitor);
        
        return { success: true, metoda: metoda, detalii: detalii, tip: tipPlatitor };
      }
    }
    
    return { success: false, error: 'Email negasit in tabel' };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}


// ═════════════════════════════════════════════════════════
// 9. FACTURARE PDF
// ═════════════════════════════════════════════════════════

function genereazaSiTrimiteFactura(email, nume, pret, valuta, numarFactura, tipFactura) {
  try {
    var templateId;
    
    if (tipFactura === 'curs1') templateId = TEMPLATE_CURS1_ID;
    else if (tipFactura === 'curs2') templateId = TEMPLATE_CURS2_ID;
    else if (tipFactura === 'pachet') templateId = TEMPLATE_PACHET_ID;
    else throw new Error('Tip factura necunoscut: ' + tipFactura);
    
    var templateDoc = DriveApp.getFileById(templateId);
    var copieDoc = templateDoc.makeCopy('Factura_' + numarFactura);
    var doc = DocumentApp.openById(copieDoc.getId());
    var body = doc.getBody();
    
    body.replaceText('{{numar_factura}}', numarFactura);
    body.replaceText('{{data}}', Utilities.formatDate(new Date(), 'Europe/Bucharest', 'dd.MM.yyyy'));
    body.replaceText('{{nume_client}}', nume);
    body.replaceText('{{email_client}}', email);
    body.replaceText('{{pret}}', pret.toFixed(2));
    body.replaceText('{{valuta}}', valuta);
    
    doc.saveAndClose();
    Utilities.sleep(2000);
    
    var pdf = copieDoc.getAs('application/pdf');
    pdf.setName('Factura_' + numarFactura + '.pdf');
    
    MailApp.sendEmail({
      to: email,
      subject: 'Factura Fiscală - Echilibru Digital',
      body: 'Bună ' + nume + ',\n\nÎți mulțumim pentru achiziție!\n\nAtașat regăsești factura fiscală aferentă comenzii tale.\n\nDatele de acces la platformă au fost trimise într-un email separat.\n\nCu respect,\nEchipa Echilibru Digital',
      attachments: [pdf],
      name: 'Echilibru Digital'
    });
    
    DriveApp.getFileById(copieDoc.getId()).setTrashed(true);
    Logger.log('[OK] Factura trimisa: ' + numarFactura);
    
  } catch(error) {
    Logger.log('[EROARE FACTURA] ' + error.toString());
    throw error;
  }
}



// ═════════════════════════════════════════════════════════
// 10. EMAILURI ACCES CURSURI
// ═════════════════════════════════════════════════════════

function trimiteEmailAccesCurs1(email, nume, parola, clientExistent, tipAcces, dataExpirare) {
  var mesaj;
  var subiect;
  
  if (tipAcces === 'Trial') {
    var dataExpirareFormatata = dataExpirare ? Utilities.formatDate(dataExpirare, 'GMT+2', 'dd.MM.yyyy HH:mm') : '7 zile de la activare';
    
    subiect = 'Acces GRATUIT 7 Zile la Cursul Tău! 🎁';
    
    if (clientExistent) {
      mesaj = 'Bună ' + nume + ',\n\nAi primit acces GRATUIT timp de 7 zile la cursul "Echilibru Digital"!\n\nDatele tale de acces (ACELAȘI cont pe care îl ai deja):\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\nAcces valabil până la: ' + dataExpirareFormatata + '\n------------------------------------\n\nIntră în platformă: https://echilibrudigital.ro/accesare\n\nDashboard Recomandări (acces PERMANENT):\nhttps://echilibrudigital.ro/recomanda-cursul\n\nDupă 7 zile, vei putea:\n- Cumpăra cursul pentru acces NELIMITAT\n- Folosi Dashboard-ul de Recomandări pentru a câștiga comisioane (50 RON/recomandare)\n\nMult succes!\nEchipa Echilibru Digital';
    } else {
      mesaj = 'Bună ' + nume + ',\n\nAi primit acces GRATUIT timp de 7 zile la cursul "Echilibru Digital"!\n\nDatele tale de acces:\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\nAcces valabil până la: ' + dataExpirareFormatata + '\n------------------------------------\n\nIntră în platformă: https://echilibrudigital.ro/accesare\n\nDashboard Recomandări (acces PERMANENT):\nhttps://echilibrudigital.ro/recomanda-cursul\n\nDupă 7 zile, vei putea:\n- Cumpăra cursul pentru acces NELIMITAT\n- Folosi Dashboard-ul de Recomandări pentru a câștiga comisioane (50 RON/recomandare)\n\nMult succes!\nEchipa Echilibru Digital';
    }
  } else {
    subiect = 'Bine Ai Venit! Acces NELIMITAT la Cursul Tău 🚀';
    
    if (clientExistent) {
      mesaj = 'Bună ' + nume + ',\n\nÎți mulțumim pentru achiziția cursului "Echilibru Digital"!\n\nAi acces NELIMITAT la curs cu ACELAȘI cont pe care îl ai deja:\n\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\n------------------------------------\n\nAccesează cursul: https://echilibrudigital.ro/accesare\nDashboard Recomandări: https://echilibrudigital.ro/recomanda-cursul\n\nVei primi factura fiscală într-un email separat.\n\nMult succes!\nEchipa Echilibru Digital';
    } else {
      mesaj = 'Bună ' + nume + ',\n\nÎți mulțumim pentru achiziție!\n\nDatele tale de acces NELIMITAT:\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\n------------------------------------\n\nAccesează cursul: https://echilibrudigital.ro/accesare\nDashboard Recomandări: https://echilibrudigital.ro/recomanda-cursul\n\nVei primi factura fiscală într-un email separat.\n\nMult succes!\nEchipa Echilibru Digital';
    }
  }
  
  MailApp.sendEmail({
    to: email,
    subject: subiect,
    body: mesaj,
    name: 'Echilibru Digital'
  });
}


function trimiteEmailAccesCurs2(email, nume, parola, clientExistent, tipAcces, dataExpirare) {
  var mesaj;
  var subiect;
  
  if (tipAcces === 'Trial') {
    var dataExpirareFormatata = dataExpirare ? Utilities.formatDate(dataExpirare, 'GMT+2', 'dd.MM.yyyy HH:mm') : '7 zile de la activare';
    
    subiect = 'Acces GRATUIT 7 Zile la Cursul Tău! 🎁';
    
    if (clientExistent) {
      mesaj = 'Bună ' + nume + ',\n\nAi primit acces GRATUIT timp de 7 zile la "Ghidul Părintelui în Era Digitală"!\n\nDatele tale de acces (ACELAȘI cont pe care îl ai deja):\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\nAcces valabil până la: ' + dataExpirareFormatata + '\n------------------------------------\n\nIntră în platformă: https://echilibrudigital.ro/accesare\n\nDashboard Recomandări (acces PERMANENT):\nhttps://echilibrudigital.ro/dashboard-recomandari\n\nDupă 7 zile, vei putea:\n- Cumpăra cursul pentru acces NELIMITAT\n- Folosi Dashboard-ul de Recomandări pentru a câștiga comisioane (50 RON/recomandare)\n\nMult succes!\nEchipa Echilibru Digital';
    } else {
      mesaj = 'Bună ' + nume + ',\n\nAi primit acces GRATUIT timp de 7 zile la "Ghidul Părintelui în Era Digitală"!\n\nDatele tale de acces:\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\nAcces valabil până la: ' + dataExpirareFormatata + '\n------------------------------------\n\nIntră în platformă: https://echilibrudigital.ro/accesare\n\nDashboard Recomandări (acces PERMANENT):\nhttps://echilibrudigital.ro/dashboard-recomandari\n\nDupă 7 zile, vei putea:\n- Cumpăra cursul pentru acces NELIMITAT\n- Folosi Dashboard-ul de Recomandări pentru a câștiga comisioane (50 RON/recomandare)\n\nMult succes!\nEchipa Echilibru Digital';
    }
  } else {
    subiect = 'Bine Ai Venit! Acces NELIMITAT la Cursul Tău 🚀';
    
    if (clientExistent) {
      mesaj = 'Bună ' + nume + ',\n\nÎți mulțumim pentru achiziția cursului pentru părinți!\n\nAi acces NELIMITAT la curs cu ACELAȘI cont pe care îl ai deja:\n\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\n------------------------------------\n\nAccesează cursul: https://echilibrudigital.ro/accesare\nDashboard Recomandări: https://echilibrudigital.ro/dashboard-recomandari\n\nVei primi factura fiscală într-un email separat.\n\nMult succes!\nEchipa Echilibru Digital';
    } else {
      mesaj = 'Bună ' + nume + ',\n\nÎți mulțumim pentru achiziție!\n\nDatele tale de acces NELIMITAT:\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\n------------------------------------\n\nAccesează cursul: https://echilibrudigital.ro/accesare\nDashboard Recomandări: https://echilibrudigital.ro/dashboard-recomandari\n\nVei primi factura fiscală într-un email separat.\n\nMult succes!\nEchipa Echilibru Digital';
    }
  }
  
  MailApp.sendEmail({
    to: email,
    subject: subiect,
    body: mesaj,
    name: 'Echilibru Digital'
  });
}


function trimiteEmailAccesPachet(email, nume, parola, clientExistent) {
  var mesaj;
  
  if (clientExistent) {
    mesaj = 'Bună ' + nume + ',\n\nÎți mulțumim pentru achiziția pachetului complet!\n\nAi acces la AMBELE cursuri cu ACELAȘI cont pe care îl ai deja:\n\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\n------------------------------------\n\nAccesează cursurile: https://echilibrudigital.ro/accesare\n\n- Curs 1: Echilibru Digital\n- Curs 2: Echilibru Digital pentru Părinți\n\nVei primi factura fiscală într-un email separat.\n\nMult succes!\nEchipa Echilibru Digital';
  } else {
    mesaj = 'Bună ' + nume + ',\n\nÎți mulțumim pentru achiziția pachetului complet!\n\nDatele tale de acces:\n------------------------------------\nEmail: ' + email + '\nParolă: ' + parola + '\n------------------------------------\n\nAccesează cursurile: https://echilibrudigital.ro/accesare\n\n- Curs 1: Echilibru Digital\n- Curs 2: Echilibru Digital pentru Părinți\n\nVei primi factura fiscală într-un email separat.\n\nMult succes!\nEchipa Echilibru Digital';
  }
  
  MailApp.sendEmail({
    to: email,
    subject: 'Bine Ai Venit! Acces la Pachetul Tău Complet 💎',
    body: mesaj,
    name: 'Echilibru Digital'
  });
}


// ═════════════════════════════════════════════════════════
// 11. API WEB (doGet) - Router principal
// ═════════════════════════════════════════════════════════

function doGet(e) {
    try {
        if (e && e.parameter && e.parameter.action === 'genereazaReferral') {
            var email = e.parameter.email;
            if (!email) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipseste parametrul email'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = genereazaLinkReferral(email);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'statisticiReferral') {
            var email = e.parameter.email;
            if (!email) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipseste parametrul email'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = getStatisticiReferral(email);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'salveazaMetodaPlata') {
            var email = e.parameter.email;
            var metoda = e.parameter.metoda;
            var detalii = e.parameter.detalii;
            var cnp = e.parameter.cnp;
            var adresa = e.parameter.adresa;
            var firma = e.parameter.firma;
            var cui = e.parameter.cui;
            
            if (!email || !metoda || !detalii) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipsesc parametri: email, metoda sau detalii'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = salveazaMetodaPlata(email, metoda, detalii, cnp, adresa, firma, cui);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'proceseazaSesiuneStripe') {
            var sessionId = e.parameter.session_id;
            var emailParam = e.parameter.email;
            if (sessionId) {
                var rezultat = proceseazaSesiuneStripe(sessionId);
                return ContentService.createTextOutput(JSON.stringify(rezultat)).setMimeType(ContentService.MimeType.JSON);
            } else if (emailParam) {
                var rezultat = proceseazaSesiuneDupaEmail(emailParam);
                return ContentService.createTextOutput(JSON.stringify(rezultat)).setMimeType(ContentService.MimeType.JSON);
            } else {
                return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Lipseste session_id sau email' }))
                    .setMimeType(ContentService.MimeType.JSON);
            }
        }
        if (e && e.parameter && e.parameter.action === 'getUltimele3Plati') {
            var rezultat = getUltimele3Plati();
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'verificaAccesCurs') {
            var email = e.parameter.email;
            var produs = e.parameter.produs;
            if (!email || !produs) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipseste email sau produs'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = verificaAccesCurs(email, produs);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'verificaAcces') {
            var email = e.parameter.email;
            var parola = e.parameter.parola;
            if (!email || !parola) {
                return ContentService.createTextOutput(JSON.stringify({
                    acces: false, error: 'Lipseste email sau parola'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
            if (!sheet) {
                return ContentService.createTextOutput(JSON.stringify({
                    acces: false, error: 'Baza de date nu este disponibila'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var data = sheet.getDataRange().getValues();
            var headers = data[0];
            var colEmail = headers.indexOf('Email');
            var colParola = headers.indexOf('Parola');
            if (colParola === -1) colParola = headers.indexOf('Parolă');
            var colNume = headers.indexOf('Nume');
            var colCurs1 = headers.indexOf('Acces Curs 1');
            var colCurs2 = headers.indexOf('Acces Curs 2');
            var colPachet = headers.indexOf('Acces Pachet');
            for (var i = 1; i < data.length; i++) {
                if (data[i][colEmail] === email && data[i][colParola] === parola) {
                    return ContentService.createTextOutput(JSON.stringify({
                        acces: true,
                        nume: data[i][colNume] || 'Client',
                        acces_curs1: data[i][colCurs1] === true || data[i][colCurs1] === 'TRUE',
                        acces_curs2: data[i][colCurs2] === true || data[i][colCurs2] === 'TRUE',
                        acces_pachet: data[i][colPachet] === true || data[i][colPachet] === 'TRUE'
                    })).setMimeType(ContentService.MimeType.JSON);
                }
            }
            return ContentService.createTextOutput(JSON.stringify({
                acces: false, error: 'Email sau parola incorecta'
            })).setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'verificaToken') {
            var token = e.parameter.token;
            if (!token) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipseste parametrul token'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = verificaAccesToken(token);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'solicitaCod') {
            var email = e.parameter.email;
            if (!email) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipseste parametrul email'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = solicitaCodVerificare(email);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'verificaCod') {
            var email = e.parameter.email;
            var cod = e.parameter.cod;
            if (!email || !cod) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipsesc parametrii email sau cod'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = verificaCodVerificare(email, cod);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        if (e && e.parameter && e.parameter.action === 'inregistrareTrial') {
            var email = e.parameter.email;
            var nume = e.parameter.nume;
            var produs = e.parameter.produs || 'curs1';
            if (!email) {
                return ContentService.createTextOutput(JSON.stringify({
                    success: false, error: 'Lipseste adresa de email.'
                })).setMimeType(ContentService.MimeType.JSON);
            }
            var rezultat = inregistrareTrialManual(email, nume, produs);
            return ContentService.createTextOutput(JSON.stringify(rezultat))
                .setMimeType(ContentService.MimeType.JSON);
        }
        return ContentService.createTextOutput('Apps Script API activ.')
            .setMimeType(ContentService.MimeType.TEXT);
    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
            error: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}



// ═════════════════════════════════════════════════════════
// 12. PROCESARE INSTANTĂ STRIPE
// ═════════════════════════════════════════════════════════

function proceseazaSesiuneStripe(sessionId) {
  try {
    Logger.log('[INSTANT] Cerere procesare pentru: ' + sessionId);
    
    var session = stripeApiRequest('checkout/sessions/' + sessionId, 'get', {});
    
    if (!session || session.status !== 'complete') {
      return { success: false, error: 'Sesiune neplatita sau invalida.' };
    }
    
    var email = session.customer_details?.email || 'necunoscut@email.com';
    var nume = session.customer_details?.name || 'Client';
    var tipProdus = (session.metadata || {}).tip_produs;
    var pret = session.amount_total / 100;
    var valuta = (session.currency || 'ron').toUpperCase();

    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (e) {
      return { success: false, error: 'Server ocupat. Reincearca in cateva secunde.' };
    }
    
    try {
      var dejaProcesata = verificaDuplicateSession(sessionId);
      var client = cautaClientExistent(email);
      
      if (dejaProcesata && client) {
        Logger.log('[INSTANT] Sesiune deja procesata. Generez token nou.');
        
        var tokenNou = genereazaTokenConfirmare();
        var tipAcces = (session.metadata || {}).tip_cupon === 'acces_gratuit' ? 'Trial' : 'Platit';
        
        salveazaToken(tokenNou, email, sessionId, client.parola, tipAcces);
        
        return {
          success: true,
          token: tokenNou,
          email: email,
          parola: client.parola,
          tipAcces: tipAcces
        };
      }
      
      marcheazaSessionProcesat(sessionId);
      lock.releaseLock();
      
      var rezultat;
      if (tipProdus === 'echilibru_digital') {
        rezultat = procesareCurs1(email, nume, pret, valuta, sessionId, session.metadata);
      } else if (tipProdus === 'curs_parinti') {
        rezultat = procesareCurs2(email, nume, pret, valuta, sessionId, session.metadata);
      } else if (tipProdus === 'pachet_redus') {
        rezultat = procesarePachetComplet(email, nume, pret, valuta, sessionId, session.metadata);
      } else {
        return { success: false, error: 'Produs necunoscut in metadata: ' + tipProdus };
      }
      
      if (rezultat.success) {
        var token = genereazaTokenConfirmare();
        salveazaToken(token, email, sessionId, rezultat.parola, rezultat.tipAcces);
        
        var sheetTokens = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Confirmation Tokens');
        if (sheetTokens) {
          var dataTokens = sheetTokens.getDataRange().getValues();
          for (var j = 1; j < dataTokens.length; j++) {
            if (dataTokens[j][0] === token) {
              sheetTokens.getRange(j + 1, 7).setValue(true);
              sheetTokens.getRange(j + 1, 8).setValue(new Date());
              break;
            }
          }
        }
        
        rezultat.token = token;
      }
      
      return rezultat;

    } catch (innerError) {
      if (lock.hasLock()) lock.releaseLock();
      throw innerError;
    }
    
  } catch (error) {
    Logger.log('[EROARE INSTANT] ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function proceseazaSesiuneDupaEmail(email) {
  try {
    Logger.log('[INSTANT] Cautare plata recenta pentru: ' + email);
    
    var sessions = stripeApiRequest('checkout/sessions', 'get', {
      limit: 15,
      status: 'complete'
    });
    
    var sesiuneGasita = null;
    if (sessions && sessions.data) {
        for (var i = 0; i < sessions.data.length; i++) {
            var s = sessions.data[i];
            if (s.customer_details && s.customer_details.email && s.customer_details.email.toLowerCase() === email.toLowerCase()) {
                if (s.metadata && s.metadata.tip_produs) {
                   sesiuneGasita = s;
                   break;
                }
            }
        }
    }
    
    if (!sesiuneGasita) {
      return { success: false, error: 'Nu am gasit nicio plata recenta pentru cursuri pe acest email: ' + email };
    }
    
    Logger.log('[INSTANT] Sesiune gasita dupa email: ' + sesiuneGasita.id);
    return proceseazaSesiuneStripe(sesiuneGasita.id);
    
  } catch (error) {
    Logger.log('[EROARE] proceseazaSesiuneDupaEmail: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}


// ═════════════════════════════════════════════════════════
// 13. SISTEM SECURITATE (Token One-Time + Cod Verificare Email)
// ═════════════════════════════════════════════════════════

function genereazaTokenConfirmare() {
  var caractere = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var token = 'conf_';
  for (var i = 0; i < 16; i++) {
    token += caractere.charAt(Math.floor(Math.random() * caractere.length));
  }
  return token;
}

function salveazaToken(token, email, sessionId, parola, tipAcces) {
  var sheet = getOrCreateSheet('Confirmation Tokens', [
    'Token', 'Email', 'Session ID', 'Parola', 'Tip Acces',
    'Data Creare', 'Folosit', 'Data Folosit'
  ]);
  
  sheet.appendRow([
    token, email, sessionId, parola, tipAcces,
    new Date(), false, ''
  ]);
}

function verificaToken(token) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Confirmation Tokens');
  if (!sheet) {
    return { success: false, error: 'Tabelul de token-uri nu exista' };
  }
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      var folosit = data[i][6];
      
      if (folosit === true || folosit === 'TRUE') {
        return { success: false, error: 'Link expirat sau deja folosit' };
      }
      
      sheet.getRange(i + 1, 7).setValue(true);
      sheet.getRange(i + 1, 8).setValue(new Date());
      
      return {
        success: true,
        email: data[i][1],
        parola: data[i][3],
        tipAcces: data[i][4]
      };
    }
  }
  
  return { success: false, error: 'Token invalid' };
}

function verificaAccesToken(token) {
  return verificaToken(token);
}

function genereazaCodVerificare() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

function salveazaCodVerificare(email, cod) {
  var sheet = getOrCreateSheet('Verification Codes', [
    'Email', 'Cod', 'Data Creare', 'Expirat', 'Folosit'
  ]);
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][3] === false) {
      sheet.getRange(i + 1, 4).setValue(true);
    }
  }
  
  sheet.appendRow([email, cod, new Date(), false, false]);
}

function verificaCodVerificare(email, cod) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Verification Codes');
  if (!sheet) {
    return { success: false, error: 'Sistem indisponibil' };
  }
  
  var data = sheet.getDataRange().getValues();
  var acum = new Date();
  var codCautat = cod.toString().trim();
  
  for (var i = data.length - 1; i >= 1; i--) {
    var emailSheet = data[i][0];
    var codSheet = data[i][1].toString().trim();
    
    if (emailSheet === email && codSheet === codCautat) {
      var dataCreare = new Date(data[i][2]);
      var expirat = data[i][3];
      var folosit = data[i][4];
      
      var diferentaMinute = (acum - dataCreare) / (1000 * 60);
      
      if (diferentaMinute > 10) {
        return { success: false, error: 'Codul a expirat. Solicita unul nou.' };
      }
      
      if (expirat === true || expirat === 'TRUE') {
        return { success: false, error: 'Codul a expirat' };
      }
      
      if (folosit === true || folosit === 'TRUE') {
        return { success: false, error: 'Cod deja folosit' };
      }
      
      sheet.getRange(i + 1, 5).setValue(true);
      
      var client = cautaClientExistent(email);
      if (!client) {
        return { success: false, error: 'Client negasit' };
      }
      
      return {
        success: true,
        email: email,
        parola: client.parola
      };
    }
  }
  
  return { success: false, error: 'Cod incorect' };
}

function trimiteEmailCodVerificare(email, cod) {
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Codul tau de verificare - Echilibru Digital',
      body: 'Buna,\n\nAi solicitat accesul la datele contului tau.\n\nCodul tau de verificare:\n\n    ' + cod + '\n\nAcest cod expira in 10 minute.\n\nDaca nu ai solicitat tu acest cod, ignora acest email.\n\nCu respect,\nEchipa Echilibru Digital',
      name: 'Echilibru Digital'
    });
    
    return { success: true };
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

function solicitaCodVerificare(email) {
  try {
    var client = cautaClientExistent(email);
    if (!client) {
      return {
        success: false,
        error: 'Nu am gasit nicio achizitie pentru acest email'
      };
    }
    
    var cod = genereazaCodVerificare();
    salveazaCodVerificare(email, cod);
    
    var rezultatEmail = trimiteEmailCodVerificare(email, cod);
    
    if (!rezultatEmail.success) {
      return {
        success: false,
        error: 'Eroare trimitere email: ' + rezultatEmail.error
      };
    }
    
    return {
      success: true,
      mesaj: 'Cod trimis pe email'
    };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}


// ═════════════════════════════════════════════════════════
// 14. SISTEM TRIAL 7 ZILE
// ═════════════════════════════════════════════════════════

function detectareCuponTrial(sessionId) {
  try {
    var session = stripeApiRequest('checkout/sessions/' + sessionId, 'get', {
      expand: ['total_details.breakdown']
    });
    
    var esteTrial = false;
    var cuponAplicat = null;
    
    if (session.total_details && session.total_details.amount_discount > 0) {
      Logger.log('[TRIAL] Discount detectat pe sesiune');
      
      if (session.total_details.breakdown && session.total_details.breakdown.discounts) {
        var discounts = session.total_details.breakdown.discounts;
        
        for (var i = 0; i < discounts.length; i++) {
          if (discounts[i].discount && discounts[i].discount.source && discounts[i].discount.source.coupon) {
            cuponAplicat = discounts[i].discount.source.coupon;
            
            if (cuponAplicat === 'acces_gratuit' || cuponAplicat.toLowerCase().includes('gratuit')) {
              esteTrial = true;
              break;
            }
          }
        }
      }
      
      if (!esteTrial && session.invoice) {
        var invoice = stripeApiRequest('invoices/' + session.invoice, 'get', {});
        
        if (invoice.discount && invoice.discount.coupon) {
          cuponAplicat = invoice.discount.coupon.id;
          
          if (invoice.discount.coupon.metadata && 
              invoice.discount.coupon.metadata.tip_cupon === 'acces_gratuit') {
            esteTrial = true;
          }
          
          if (cuponAplicat === 'acces_gratuit' || cuponAplicat.toLowerCase().includes('gratuit')) {
            esteTrial = true;
          }
        }
      }
    }
    
    var tipAcces = esteTrial ? 'Trial' : 'Platit';
    
    return {
      esteTrial: esteTrial,
      cuponId: cuponAplicat,
      tipAcces: tipAcces,
      dataExpirare: esteTrial ? new Date(Date.now() + 7 * 24 * 60 * 60 * 1000) : null
    };
    
  } catch(error) {
    Logger.log('[EROARE TRIAL] detectareCupon: ' + error.toString());
    return {
      esteTrial: false,
      cuponId: null,
      tipAcces: 'Platit',
      dataExpirare: null
    };
  }
}

function salveazaTrial(email, produs, sursa, dataInscriere, dataExpirare) {
  var sheet = getOrCreateSheet('Trial Access', [
    'Email', 'Produs', 'Cupon Aplicat', 'Data Start', 'Data Expirare',
    'Status', 'Data Conversie Plata', 'Tip Conversie',
    'Email Ziua 4 Trimis', 'Email Ziua 7 Trimis'
  ]);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1] === produs) {
      return;
    }
  }
  sheet.appendRow([
    email, produs, sursa, dataInscriere, dataExpirare,
    'Activ', '', '', '', ''
  ]);
}

function actualizeazaTrialLaConvertit(email, produs) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Trial Access');
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][1] === produs && data[i][5] === 'Activ') {
        sheet.getRange(i + 1, 6).setValue('Convertit');
        sheet.getRange(i + 1, 7).setValue(new Date());
        sheet.getRange(i + 1, 8).setValue('platit');
        break;
      }
    }
  } catch(error) {
    Logger.log('[EROARE] actualizeazaTrialLaConvertit: ' + error.toString());
  }
}

function verificaAccesCurs(email, produs) {
  try {
    var clientSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
    if (!clientSheet) {
      return { success: false, error: 'Tabelul Clienti nu exista' };
    }
    
    var data = clientSheet.getDataRange().getValues();
    
    var mapareProduseColoane = {
      'curs1': 4,
      'curs_parinti': 5,
      'pachet': 6
    };
    
    var mapareProduseURL = {
      'curs1': '/testul-sincer',
      'curs_parinti': '/lectia1'
    };
    
    var coloanaAcces = mapareProduseColoane[produs];
    if (!coloanaAcces) {
      return { success: false, error: 'Produs invalid: ' + produs };
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        var areAcces = data[i][coloanaAcces] === true || data[i][coloanaAcces] === 'TRUE';
        
        if (!areAcces) {
          var accesCurs1 = data[i][4] === true || data[i][4] === 'TRUE';
          var accesCurs2 = data[i][5] === true || data[i][5] === 'TRUE';
          var tipAcces = data[i][9] || '';
          var dataExpirare = data[i][10];
          
          if (tipAcces === 'Trial' && dataExpirare) {
            var acum = new Date();
            var expirare = new Date(dataExpirare);
            
            if (acum < expirare) {
              var cursDisponibil = null;
              if (produs === 'curs1' && accesCurs2) {
                cursDisponibil = mapareProduseURL['curs_parinti'];
              } else if (produs === 'curs_parinti' && accesCurs1) {
                cursDisponibil = mapareProduseURL['curs1'];
              }
              
              return {
                areAcces: false,
                motiv: 'trial_activ_alt_curs',
                cursDisponibil: cursDisponibil,
                tipAcces: 'trial',
                dataExpirare: Utilities.formatDate(expirare, 'GMT+2', 'dd.MM.yyyy HH:mm'),
                zileRamase: Math.ceil((expirare - acum) / (1000 * 60 * 60 * 24))
              };
            }
          }
          
          if (tipAcces === 'Trial Expirat' || (tipAcces === 'Trial' && dataExpirare && new Date() >= new Date(dataExpirare))) {
            return {
              areAcces: false,
              motiv: 'trial_expirat',
              cursDisponibil: null,
              tipAcces: 'expirat',
              dataExpirare: dataExpirare ? Utilities.formatDate(new Date(dataExpirare), 'GMT+2', 'dd.MM.yyyy HH:mm') : null,
              zileRamase: 0
            };
          }
          
          return {
            areAcces: false,
            motiv: 'nu_achizitionat',
            cursDisponibil: null,
            tipAcces: 'fara_acces',
            dataExpirare: null,
            zileRamase: null
          };
        }
        
        var tipAcces = data[i][9] || 'Platit';
        var dataExpirare = data[i][10];
        
        if (tipAcces === 'Platit') {
          return {
            areAcces: true,
            motiv: 'acces_platit',
            tipAcces: 'platit',
            dataExpirare: null,
            zileRamase: null
          };
        }
        
        if (tipAcces === 'Trial' && dataExpirare) {
          var acum = new Date();
          var expirare = new Date(dataExpirare);
          
          if (acum < expirare) {
            var zileRamase = Math.ceil((expirare - acum) / (1000 * 60 * 60 * 24));
            return {
              areAcces: true,
              motiv: 'trial_activ',
              tipAcces: 'trial',
              dataExpirare: Utilities.formatDate(expirare, 'GMT+2', 'dd.MM.yyyy HH:mm'),
              expirareTimestamp: expirare.getTime(),
              zileRamase: zileRamase
            };
          } else {
            return {
              areAcces: false,
              motiv: 'trial_expirat',
              cursDisponibil: null,
              tipAcces: 'expirat',
              dataExpirare: Utilities.formatDate(expirare, 'GMT+2', 'dd.MM.yyyy HH:mm'),
              zileRamase: 0
            };
          }
        }
        
        if (tipAcces === 'Trial Expirat') {
          return {
            areAcces: false,
            motiv: 'trial_expirat',
            cursDisponibil: null,
            tipAcces: 'expirat',
            dataExpirare: dataExpirare ? Utilities.formatDate(new Date(dataExpirare), 'GMT+2', 'dd.MM.yyyy HH:mm') : null,
            zileRamase: 0
          };
        }
        
        return {
          areAcces: true,
          motiv: 'acces_platit',
          tipAcces: 'platit',
          dataExpirare: null,
          zileRamase: null
        };
      }
    }
    
    return {
      areAcces: false,
      motiv: 'nu_achizitionat',
      cursDisponibil: null,
      tipAcces: 'fara_acces',
      dataExpirare: null,
      zileRamase: null
    };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}



// ═════════════════════════════════════════════════════════
// 15. INREGISTRARE TRIAL MANUAL (din formular extern)
// ═════════════════════════════════════════════════════════

function inregistrareTrialManual(email, nume, produs) {
    try {
        Logger.log('[TRIAL MANUAL] Inregistrare pentru: ' + email + ' la produsul: ' + produs);
        var mapareColoane = { 'curs1': 5, 'curs_parinti': 6, 'curs2': 6 };
        var coloanaAcces = mapareColoane[produs] || 5;
        
        var client = cautaClientExistent(email);
        var parola;
        var esteClientNou = false;
        var dataExpirareFinala;
        
        if (client) {
            parola = client.parola;
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
            if (sheet) {
                sheet.getRange(client.rand, coloanaAcces).setValue(true);
                sheet.getRange(client.rand, 10).setValue('Trial');
                
                var cellExpirare = sheet.getRange(client.rand, 11);
                var dataExpirareExistenta = cellExpirare.getValue();
                if (dataExpirareExistenta && dataExpirareExistenta instanceof Date && !isNaN(dataExpirareExistenta)) {
                    dataExpirareFinala = dataExpirareExistenta;
                } else {
                    dataExpirareFinala = new Date();
                    dataExpirareFinala.setDate(dataExpirareFinala.getDate() + 7);
                    cellExpirare.setValue(dataExpirareFinala);
                }
            }
        } else {
            esteClientNou = true;
            parola = genereazaParola();
            nume = nume || 'Client Nou';
            var sheetClienti = getOrCreateSheet('Clienti', ['Email', 'Nume', 'Data Inscriere', 'Parola', 'Acces Curs 1', 'Acces Curs 2', 'Acces Pachet', 'Data ultima achizitie', 'Total Achizitii', 'Tip Acces', 'Data Expirare Trial']);
            dataExpirareFinala = new Date();
            dataExpirareFinala.setDate(dataExpirareFinala.getDate() + 7);
            var newRow = [
                email, nume, new Date(), parola,
                false, false, false, '', '', 'Trial', dataExpirareFinala
            ];
            if (produs === 'curs1') newRow[4] = true;
            if (produs === 'curs_parinti' || produs === 'curs2') newRow[5] = true;
            sheetClienti.appendRow(newRow);
        }
        
        salveazaTrial(email, produs, 'formular_gratuit', new Date(), dataExpirareFinala);
        
        try {
            if (produs === 'curs_parinti' || produs === 'curs2') {
                if (typeof trimiteEmailAccesCurs2 === 'function') {
                    trimiteEmailAccesCurs2(email, nume || 'Client', parola, !esteClientNou, 'Trial', dataExpirareFinala);
                }
            } else {
                if (typeof trimiteEmailAccesCurs1 === 'function') {
                    trimiteEmailAccesCurs1(email, nume || 'Client', parola, !esteClientNou, 'Trial', dataExpirareFinala);
                }
            }
        } catch (e) {
            Logger.log('[EROARE] Email welcome trial: ' + e.toString());
        }
        
        var token = genereazaTokenConfirmare();
        var sessionId = 'manual_' + new Date().getTime();
        salveazaToken(token, email, sessionId, parola, 'Trial');
        
        return { success: true, token: token, mesaj: 'Cont creat si acces activat.' };
    } catch (error) {
        Logger.log('[EROARE] inregistrareTrialManual: ' + error.toString());
        return { success: false, error: error.toString() };
    }
}


// ═════════════════════════════════════════════════════════
// 16. AUTOMATIZARE EMAILURI TRIAL (Ziua 4 si Ziua 7)
// ═════════════════════════════════════════════════════════

function trimiteEmailuriAutomateTrial() {
  try {
    var sheet = getOrCreateSheet('Trial Access', [
      'Email', 'Produs', 'Cupon Aplicat', 'Data Start', 
      'Data Expirare', 'Status', 'Data Conversie Platita', 'Tip Conversie',
      'Email Ziua 4 Trimis', 'Email Ziua 7 Trimis'
    ]);
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colEmail = headers.indexOf('Email');
    var colDataStart = headers.indexOf('Data Start');
    var colStatus = headers.indexOf('Status');
    var colZiua4 = headers.indexOf('Email Ziua 4 Trimis');
    var colZiua7 = headers.indexOf('Email Ziua 7 Trimis');
    var colProdus = headers.indexOf('Produs');
    
    if (colZiua4 === -1) {
      colZiua4 = headers.length;
      sheet.getRange(1, colZiua4 + 1).setValue('Email Ziua 4 Trimis');
    }
    if (colZiua7 === -1) {
      colZiua7 = headers.length + (colZiua4 === headers.length ? 1 : 0);
      sheet.getRange(1, colZiua7 + 1).setValue('Email Ziua 7 Trimis');
    }
    
    var astazi = new Date();
    var contorZiua4 = 0;
    var contorZiua7 = 0;
    
    for (var i = 1; i < data.length; i++) {
      var email = data[i][colEmail];
      var dataStart = data[i][colDataStart];
      var status = data[i][colStatus];
      var produs = data[i][colProdus];
      
      if (!dataStart || !(dataStart instanceof Date)) continue;
      
      var diferentaMs = astazi - new Date(dataStart);
      var zileTrecute = Math.floor(diferentaMs / (1000 * 60 * 60 * 24));
      
      if (zileTrecute >= 4 && (!data[i][colZiua4] || data[i][colZiua4] === '')) {
        trimiteEmailZiua4(email);
        sheet.getRange(i + 1, colZiua4 + 1).setValue('DA - ' + Utilities.formatDate(new Date(), 'GMT+2', 'dd.MM.yyyy'));
        contorZiua4++;
      }
      
      if (zileTrecute >= 7 && (!data[i][colZiua7] || data[i][colZiua7] === '') && status !== 'Convertit') {
        trimiteEmailZiua7(email, produs);
        sheet.getRange(i + 1, colZiua7 + 1).setValue('DA - ' + Utilities.formatDate(new Date(), 'GMT+2', 'dd.MM.yyyy'));
        
        if (status === 'Activ') {
          sheet.getRange(i + 1, colStatus + 1).setValue('Expirat');
        }
        contorZiua7++;
      }
    }
    
    Logger.log('[OK] Automatizare completa. Ziua 4: ' + contorZiua4 + ', Ziua 7: ' + contorZiua7);
    
  } catch(error) {
    Logger.log('[EROARE] trimiteEmailuriAutomateTrial: ' + error.toString());
  }
}

function trimiteEmailZiua4(email) {
  var nume = extragePrenum(email);
  var subiect = '50 RON te așteaptă (încă e valabil!) 💸';
  
  var mesaj = 'Bună ' + nume + ',\n\nSper că îți place cursul gratuit până acum!\n\nVoiam doar să îți reamintesc că poți câștiga 50 RON pentru fiecare părinte pe care îl ajuți.\n\nCum funcționează:\n1. Recomanzi cursul unei prietene.\n2. Ea primește acces (exact ca tine).\n3. Dacă decide să rămână, tu primești 50 RON.\n\nEste simplu. Nu trebuie să vinzi nimic, doar să ajuți.\n\nIa-ți link-ul unic de aici:\nhttps://echilibrudigital.ro/recomanda-cursul\n\nSpor la vizionat (și la câștigat)!\n\nCu drag,\nEchipa Echilibru Digital';

  try {
    MailApp.sendEmail({
      to: email,
      subject: subiect,
      body: mesaj,
      name: 'Echilibru Digital'
    });
  } catch (e) {
    Logger.log('[EROARE] Email ziua 4: ' + e.toString());
  }
}

function trimiteEmailZiua7(email, produs) {
  var nume = extragePrenum(email);
  var subiect = 'Accesul tău gratuit a EXPIRAT (Dar am o surpriză...) 🎁';
  
  var altCurs = (produs === 'curs1') ? '"Ghidul Părintelui în Era Digitală"' : '"Echilibru Digital"';
  
  var mesaj = 'Bună ' + nume + ',\n\nCele 7 zile de acces gratuit s-au încheiat. Sperăm că informațiile te-au ajutat deja să faci o schimbare în bine!\n\nVrem să te răsplătim pentru că ai parcurs materialele. Așa că avem o propunere unică pentru tine:\n\nOFERTĂ FINALĂ (BOGO): Cumpără 1, Primești 2\n\nDacă decizi să cumperi accesul nelimitat la cursul pe care l-ai parcurs (99 RON), primești CADOU și celălalt curs al nostru, ' + altCurs + '!\n\n- Acces Nelimitat pe Viață la ambele cursuri.\n- Doar 99 RON (preț unic).\n- PLUS: Păstrezi dreptul de a câștiga 50 RON/recomandare.\n\nAceastă ofertă este modul nostru de a spune "Mulțumim" că ne ești alături.\n\nActivează Oferta 1+1 aici:\nhttps://echilibrudigital.ro/oferta-finala\n\nNu lăsa progresul făcut să se piardă. Continuă acum!\n\nCu respect,\nEchipa Echilibru Digital';

  try {
    MailApp.sendEmail({
      to: email,
      subject: subiect,
      body: mesaj,
      name: 'Echilibru Digital'
    });
  } catch (e) {
    Logger.log('[EROARE] Email ziua 7: ' + e.toString());
  }
}


// ═════════════════════════════════════════════════════════
// 17. SISTEM PLATI COMISIOANE
// ═════════════════════════════════════════════════════════

function actualizeazaStatusuriPlati() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Conversii Referral');
    
    if (!sheet) {
      Logger.log('[ATENTIE] Tabelul "Conversii Referral" nu exista!');
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var colData = -1;
    var colStatusPlata = -1;
    var colDataPlata = -1;
    
    for (var j = 0; j < headers.length; j++) {
      if (headers[j] === 'Data') colData = j;
      if (headers[j] === 'Status Plata Comision') colStatusPlata = j;
      if (headers[j] === 'Data Plata Comision') colDataPlata = j;
    }
    
    if (colData === -1) {
      Logger.log('[ATENTIE] Coloana "Data" nu a fost gasita!');
      return;
    }
    
    if (colStatusPlata === -1) {
      colStatusPlata = headers.length;
      sheet.getRange(1, colStatusPlata + 1).setValue('Status Plata Comision');
    }
    
    if (colDataPlata === -1) {
      colDataPlata = headers.length + (colStatusPlata === headers.length ? 1 : 0);
      sheet.getRange(1, colDataPlata + 1).setValue('Data Plata Comision');
    }
    
    var astazi = new Date();
    astazi.setHours(0, 0, 0, 0);
    
    var countPlatibile = 0;
    var countAsteptare = 0;
    var countPlatite = 0;
    
    for (var i = 1; i < data.length; i++) {
      var dataConversie = data[i][colData];
      var statusPlata = data[i][colStatusPlata] || '';
      var dataPlata = data[i][colDataPlata];
      
      if (dataPlata && dataPlata !== '') {
        countPlatite++;
        continue;
      }
      
      if (!dataConversie || !(dataConversie instanceof Date)) {
        continue;
      }
      
      var dataConv = new Date(dataConversie);
      dataConv.setHours(0, 0, 0, 0);
      
      var diferentaMs = astazi - dataConv;
      var zileTrecute = Math.floor(diferentaMs / (1000 * 60 * 60 * 24));
      
      var statusNou = '';
      var culoare = '';
      
      if (zileTrecute >= 14) {
        statusNou = 'Poate fi platit';
        culoare = '#d9ead3';
        countPlatibile++;
      } else {
        var zileRamase = 14 - zileTrecute;
        statusNou = 'Se poate plati in ' + zileRamase + (zileRamase === 1 ? ' zi' : ' zile');
        culoare = '#fff2cc';
        countAsteptare++;
      }
      
      if (statusNou !== statusPlata) {
        var range = sheet.getRange(i + 1, colStatusPlata + 1);
        range.setValue(statusNou);
        range.setBackground(culoare);
        range.setFontWeight('bold');
      }
    }
    
    Logger.log('[OK] Actualizare statusuri plati completa!');
    Logger.log('Pot fi platite: ' + countPlatibile + ', In asteptare: ' + countAsteptare + ', Deja platite: ' + countPlatite + ', Total: ' + (data.length - 1));
    
  } catch(error) {
    Logger.log('[EROARE] actualizeazaStatusuriPlati: ' + error.toString());
  }
}

function adaugaInIstoricPlati(emailClient, prenume, suma, dataPlata) {
  try {
    var sheet = getOrCreateSheet('Istoric Plati', [
      'Data Plata', 'Email Client', 'Prenume', 'Suma Platita'
    ]);
    
    var data = dataPlata || new Date();
    
    // MODIFICARE INTELIGENTA: Cautam primul rand gol pe baza coloanei A (Data)
    // Deoarece utilizatorul a umplut coloana E cu checkbox-uri, appendRow considera tabelul plin pana jos.
    
    var lastRow = sheet.getLastRow();
    var rangeA = sheet.getRange(1, 1, lastRow, 1).getValues(); 
    var randTinta = lastRow + 1; // Default: la final de tot
    
    for (var i = 1; i < rangeA.length; i++) {
        if (!rangeA[i][0]) { // Daca celula din col A e goala
            randTinta = i + 1;
            break;
        }
    }
    
    sheet.getRange(randTinta, 1, 1, 4).setValues([[data, emailClient, prenume, suma]]);
    
    SpreadsheetApp.flush(); // MENTINEM FLUSH PENTRU SIGURANTA
    
    return { success: true };
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

function getUltimele3Plati() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Istoric Plati');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, plati: [] };
    }
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    
    data.sort(function(a, b) {
      return new Date(b[0]) - new Date(a[0]);
    });
    
    var ultimele3 = [];
    var count = Math.min(3, data.length);
    
    for (var i = 0; i < count; i++) {
      var row = data[i];
      var datePlata = new Date(row[0]);
      var prenume = row[2] || extragePrenum(row[1]);
      var suma = row[3];
      
      var zileScurse = Math.floor((new Date() - datePlata) / (1000 * 60 * 60 * 24));
      var timpText;
      
      if (zileScurse === 0) timpText = 'astazi';
      else if (zileScurse === 1) timpText = 'acum 1 zi';
      else if (zileScurse <= 7) timpText = 'acum ' + zileScurse + ' zile';
      else timpText = 'pe ' + Utilities.formatDate(datePlata, 'GMT+2', 'dd.MM');
      
      ultimele3.push({
        prenume: prenume,
        suma: suma,
        timpText: timpText,
        dataPlata: Utilities.formatDate(datePlata, 'GMT+2', 'dd.MM.yyyy')
      });
    }
    
    return { success: true, plati: ultimele3 };
    
  } catch(error) {
    return { success: true, plati: [] };
  }
}

function marcheazaPlata(emailReferrer, emailClientNou, sumaPlata, prenume) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Conversii Referral');
    if (!sheet) {
      return { success: false, error: 'Tabelul Conversii Referral nu exista' };
    }
    
    var data = sheet.getDataRange().getValues();
    var dataPlata = new Date();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === emailReferrer && data[i][3] === emailClientNou) {
        var dataFormatata = Utilities.formatDate(dataPlata, 'GMT+2', 'dd.MM.yyyy');
        sheet.getRange(i + 1, 9).setValue('Platit la ' + dataFormatata);
        sheet.getRange(i + 1, 9).setBackground('#cce5ff');
        sheet.getRange(i + 1, 9).setFontColor('#004085');
        sheet.getRange(i + 1, 9).setFontWeight('bold');
        
        sheet.getRange(i + 1, 10).setValue(dataPlata);
        
        actualizareRonPrimiti(emailReferrer, sumaPlata);
        
        var prenumeExtras = prenume || extragePrenum(emailReferrer);
        adaugaInIstoricPlati(emailReferrer, prenumeExtras, sumaPlata, dataPlata);
        
        return { success: true, mesaj: 'Plata marcata cu succes!' };
      }
    }
    
    return { success: false, error: 'Nu s-a gasit conversia specificata' };
    
  } catch(error) {
    return { success: false, error: error.toString() };
  }
}

function actualizareRonPrimiti(emailClient, sumaNoua) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Referrals');
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === emailClient) {
        var sumaActuala = data[i][8] || 0;
        var sumaTotala = Number(sumaActuala) + Number(sumaNoua);
        sheet.getRange(i + 1, 9).setValue(sumaTotala);
        break;
      }
    }
  } catch(error) {
    Logger.log('[EROARE] actualizareRonPrimiti: ' + error);
  }
}

function onEditReferrals(e) {
  try {
    if (!e || !e.range) return;
    
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var row = e.range.getRow();
    var col = e.range.getColumn();
    
    // CAZ 1: Actualizare suma in Referrals -> Adaugare in Istoric Plati
    if (sheetName === 'Referrals' && col === 9 && row > 1) {
      // Helper simplu pentru curatare numar (scoate "RON" sau alte caractere)
      var cleanNumber = function(val) {
        if (!val) return 0;
        var s = String(val).replace(/[^0-9.-]+/g, '');
        return parseFloat(s) || 0;
      };

      var valoraNoua = cleanNumber(e.value);
      var valoraVeche = cleanNumber(e.oldValue);
      var diferenta = valoraNoua - valoraVeche;
      
      // Logica de siguranta: daca diferenta e 0 sau negativa, poate e o corectie, nu o plata noua.
      if (diferenta <= 0) return;
      
      var emailClient = sheet.getRange(row, 1).getValue();
      var prenume = extragePrenum(emailClient);
      
      var rezultat = adaugaInIstoricPlati(emailClient, prenume, diferenta, new Date());
      
      if (rezultat.success) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          prenume + ' - ' + diferenta + ' RON adaugat in Istoric Plati',
          'Plata inregistrata',
          3
        );
      }
    }
    
    // CAZ 2: Bifare "Generare PV" in Istoric Plati
    // Presupunem ca checkbox-ul este pe coloana 5 (E) - "Genereaza PV"
    // Istoric Plati headers: Data Plata, Email Client, Prenume, Suma Platita, [Genereaza PV]
    if (sheetName === 'Istoric Plati' && col === 5 && row > 1) {
      if (e.value === 'TRUE') {
        var email = sheet.getRange(row, 2).getValue();
        var suma = sheet.getRange(row, 4).getValue();
        var dataPlata = sheet.getRange(row, 1).getValue();
        
        SpreadsheetApp.getActiveSpreadsheet().toast('Se genereaza PV pentru ' + email + '...', 'Procesare', -1);
        
        var rezPV = genereazaSiTrimitePV(email, suma, dataPlata);
        
        if (rezPV.success) {
          e.range.setValue('PV Trimis: ' + rezPV.numarPV);
          SpreadsheetApp.getActiveSpreadsheet().toast('PV trimis cu succes!', 'Finalizat', 3);
        } else {
          e.range.setValue('Eroare: ' + rezPV.error);
          SpreadsheetApp.getActiveSpreadsheet().toast('Eroare generare PV: ' + rezPV.error, 'Eroare', 5);
        }
      }
    }
    
    // CAZ 3: Bifare "Cere Factura" in Referrals (Coloana 15 - O)
    // Presupunem ca utilizatorul adauga checkbox pe coloana 15 in Referrals
    if (sheetName === 'Referrals' && col === 15 && row > 1) {
      if (e.value === 'TRUE') {
        var email = sheet.getRange(row, 1).getValue();
        SpreadsheetApp.getActiveSpreadsheet().toast('Se calculeaza suma si se trimite cerere factura...', 'Procesare', -1);
        
        var rezFactura = trimiteCerereFactura(email);
        
        if (rezFactura.success) {
          e.range.setValue(false); // Debifam
          sheet.getRange(row, 16).setValue('Factura ceruta: ' + Utilities.formatDate(new Date(), 'GMT+2', 'dd.MM.yyyy') + ' (' + rezFactura.suma + ' RON)');
          SpreadsheetApp.getActiveSpreadsheet().toast('Cerere factura trimisa pentru ' + rezFactura.suma + ' RON!', 'Succes', 5);
        } else {
          e.range.setValue(false);
          SpreadsheetApp.getActiveSpreadsheet().toast('Eroare: ' + rezFactura.error, 'Eroare', 5);
        }
      }
    }
    
  } catch(error) {
    Logger.log('[EROARE] trigger onEdit: ' + error.toString());
  }
}

function trimiteCerereFactura(email) {
  try {
    // 1. Calculam suma neplatita direct din Referrals (Total Castigat - Total Primit)
    var sheetReferrals = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Referrals');
    if (!sheetReferrals) return { success: false, error: 'Nu gasesc tabelul Referrals' };
    
    var data = sheetReferrals.getDataRange().getValues();
    var sumaTotala = 0;
    var gasit = false;
    var emailCautat = String(email).trim().toLowerCase();
    
    // Cautam randul partenerului
    for (var i = 1; i < data.length; i++) {
       var refEmail = String(data[i][0] || '').trim().toLowerCase();
       
       if (refEmail === emailCautat) {
           gasit = true;
           var totalCastigat = parseFloat(data[i][5]) || 0; // Col F (6)
           var totalPrimit = parseFloat(data[i][8]) || 0;   // Col I (9)
           
           sumaTotala = totalCastigat - totalPrimit;
           break;
       }
    }
    
    if (!gasit) {
        return { success: false, error: 'Emailul partenerului nu a fost gasit in Referrals.' };
    }
    
    if (sumaTotala <= 0) {
        return { success: false, error: 'Nu exista sume restante de plata (Total Castigat: ' + totalCastigat + ', Total Primit: ' + totalPrimit + ').' };
    }
    
    // 2. Trimitem Email
    var nume = extragePrenum(email);
    var subiect = 'Emitere Factură - Comisioane Echilibru Digital';
    
    // AICI TREBUIE DATELE TALE REALE (DDA)
    var corp = 'Bună ' + nume + ',\n\n' +
               'Ai acumulat suma de ' + sumaTotala + ' RON din comisioanele de recomandare.\n\n' +
               'Te rugăm să ne emiți o factură fiscală pe datele noastre (Titular Drepturi de Autor) pentru a putea procesa plata:\n\n' +
               '--- DATE FACTURARE (DDA) ---\n' +
               'Nume: [NUMELE_TĂU]\n' +
               'CNP: [CNP-UL_TĂU]\n' +
               'Adresa: [ADRESA_TA_DIN_BULETIN]\n' +
               'Banca: [NUME_BANCĂ]\n' +
               'IBAN: [CONTUL_TĂU_IBAN]\n' +
               '---------------------------\n\n' +
               'Te rugăm să trimiți factura la acest email.\n\n' +
               'Mulțumim pentru colaborare!\n\n' +
               'Cu respect,\n' +
               'Echilibru Digital';
               
    MailApp.sendEmail({
      to: email,
      subject: subiect,
      body: corp,
      name: 'Echilibru Digital'
    });
    
    return { success: true, suma: sumaTotala };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function genereazaSiTrimitePV(email, suma, dataPlata) {
  try {
    // 1. Cautam datele extinse in Referrals
    var sheetRef = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Referrals');
    var dataRef = sheetRef.getDataRange().getValues();
    var infoPartener = null;
    
    for (var i = 1; i < dataRef.length; i++) {
      if (dataRef[i][0] === email) {
        infoPartener = {
          nume: extragePrenum(email), // Sau nume complet daca ar exista
          cnp: dataRef[i][9],
          adresa: dataRef[i][10],
          firma: dataRef[i][11],
          cui: dataRef[i][12],
          detaliiPlata: dataRef[i][7],
          tip: dataRef[i][13] // PF sau PJ
        };
        // Daca nu avem tip setat, deducem
        if (!infoPartener.tip) {
          infoPartener.tip = (infoPartener.cui && infoPartener.cui.length > 2) ? 'PJ' : 'PF';
        }
        break;
      }
    }
    
    if (!infoPartener) return { success: false, error: 'Nu am gasit datele partenerului in Referrals' };
    
    // 2. Determinam Template si Calcule
    var templateId, venitBrut, impozit, venitNet;
    
    if (infoPartener.tip === 'PJ') {
      templateId = TEMPLATE_PV_PJ_ID;
      venitBrut = suma;
      impozit = 0;
      venitNet = suma;
    } else {
      templateId = TEMPLATE_PV_PF_ID;
      
      // MODIFICARE: Calcul invers pentru a asigura ca NET-ul este fix (ex: 50 RON)
      // Daca userul primeste 50 RON in cont, asta inseamna ca suma bruta trebuie sa fie mai mare.
      // Net = Brut - (Brut * 0.1) => Net = Brut * 0.9 => Brut = Net / 0.9
      
      venitNet = Number(suma);
      venitBrut = venitNet / 0.9;
      impozit = venitBrut - venitNet;
      
      venitBrut = venitBrut.toFixed(2);
      impozit = impozit.toFixed(2);
      venitNet = venitNet.toFixed(2);
    }
    
    if (!infoPartener.detaliiPlata) infoPartener.detaliiPlata = 'Transfer Bancar / Revolut';
    var numarPV = generareNumarPV();
    var dataFormatata = Utilities.formatDate(dataPlata || new Date(), 'Europe/Bucharest', 'dd.MM.yyyy');
    
    // 3. Generare PDF
    var templateDoc = DriveApp.getFileById(templateId);
    var copieDoc = templateDoc.makeCopy('PV_' + numarPV);
    var doc = DocumentApp.openById(copieDoc.getId());
    var body = doc.getBody();
    
    body.replaceText('{{numar_pv}}', numarPV);
    body.replaceText('{{data_curenta}}', dataFormatata);
    body.replaceText('{{detalii_plata}}', infoPartener.detaliiPlata);
    body.replaceText('{{suma_bruta}}', venitBrut);
    
    if (infoPartener.tip === 'PJ') {
        body.replaceText('{{denumire_firma}}', infoPartener.firma || infoPartener.nume); // Fallback la nume daca nu e firma
        body.replaceText('{{cui_partener}}', infoPartener.cui || '-');
    } else {
        body.replaceText('{{nume_partener}}', infoPartener.nume); // Numele din email/sistem
        body.replaceText('{{cnp_partener}}', infoPartener.cnp || '-');
        body.replaceText('{{adresa_partener}}', infoPartener.adresa || '-');
        body.replaceText('{{suma_impozit}}', impozit);
        body.replaceText('{{suma_neta}}', venitNet);
    }
    
    doc.saveAndClose();
    Utilities.sleep(2000); // Asteptam propagarea
    
    var pdf = copieDoc.getAs('application/pdf');
    pdf.setName('PV_' + numarPV + '.pdf');
    
    // 4. Trimitere Email
    var subiect = 'Proces Verbal de Plată - ' + numarPV;
    var mesaj = 'Bună,\n\nAtașat regăsești Procesul Verbal pentru plata efectuată în data de ' + dataFormatata + '.\n\nSuma: ' + suma + ' RON\n\nÎți mulțumim pentru colaborare!\n\nCu respect,\nEchipa Echilibru Digital';
    
    MailApp.sendEmail({
      to: email,
      subject: subiect,
      body: mesaj,
      attachments: [pdf],
      name: 'Echilibru Digital'
    });
    
    DriveApp.getFileById(copieDoc.getId()).setTrashed(true);
    
    // Salvare in Arhiva Drive
    try {
      var folderArhiva = DriveApp.getFolderById(FOLDER_PV_ARHIVA_ID);
      folderArhiva.createFile(pdf);
    } catch(e) {
      Logger.log('[EROARE ARHIVARE PV] ' + e.toString());
      // Nu blocam executia, doar logam eroarea de salvare
    }
    
    return { success: true, numarPV: numarPV };
    
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}


// ═════════════════════════════════════════════════════════
// 18. DEBUG & TESTARE
// ═════════════════════════════════════════════════════════

function testeazaConexiuneStripe() {
  try {
    var balance = stripeApiRequest('balance', 'get', {});
    Logger.log('[OK] Conexiune Stripe OK!');
    Logger.log('Balance disponibil: ' + balance.available[0].amount / 100 + ' ' + balance.available[0].currency.toUpperCase());
    return true;
  } catch(error) {
    Logger.log('[EROARE] Conexiune Stripe: ' + error.toString());
    return false;
  }
}

function veziUltimele10Plati() {
  try {
    var sessions = stripeApiRequest('checkout/sessions', 'get', { 
      limit: 10,
      status: 'complete'
    });
    
    Logger.log('--- ULTIMELE 10 PLATI ---');
    
    sessions.data.forEach(function(s, index) {
      Logger.log((index + 1) + '. ' + s.customer_details.email + ' | ' + (s.amount_total / 100) + ' ' + s.currency.toUpperCase() + ' | ' + new Date(s.created * 1000));
    });
    
  } catch(error) {
    Logger.log('[EROARE] ' + error.toString());
  }
}

function testeazaGenerareReferral() {
  var emailTest = 'test@example.com';
  var rezultat = genereazaLinkReferral(emailTest);
  Logger.log('--- TEST GENERARE REFERRAL ---');
  Logger.log(JSON.stringify(rezultat, null, 2));
}

function testeazaAPI() {
  var emailTest = 'test@example.com';
  var parolaTest = 'TestParola123';
  
  Logger.log('[TEST] TESTARE API verificaAcces...');
  
  var mockEvent = {
    parameter: {
      action: 'verificaAcces',
      email: emailTest,
      parola: parolaTest
    }
  };
  
  var raspuns = doGet(mockEvent);
  var rezultat = JSON.parse(raspuns.getContent());
  
  Logger.log('Raspuns API: ' + JSON.stringify(rezultat, null, 2));
  
  if (rezultat.acces) {
    Logger.log('[OK] API functioneaza corect!');
    Logger.log('   Acces Curs 1: ' + rezultat.acces_curs1);
    Logger.log('   Acces Curs 2: ' + rezultat.acces_curs2);
    Logger.log('   Acces Pachet: ' + rezultat.acces_pachet);
  } else {
    Logger.log('[FAIL] API a returnat acces negat');
  }
}

function testActualizeazaStatusuriPlati() {
  actualizeazaStatusuriPlati();
  Logger.log('[OK] Test completat! Verifica tabelul "Conversii Referral"');
}

function testLogin() {
  var email = 'cozmadaniel18@gmail.com';
  var parola = 'ZZeC7wx2YW';
  
  Logger.log('--- TEST LOGIN ---');
  Logger.log('Email test: ' + email);
  Logger.log('Parola test: ' + parola);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  Logger.log('Headers gasite: ' + JSON.stringify(headers));
  
  var colEmail = headers.indexOf('Email');
  var colParola = headers.indexOf('Parola');
  var colNume = headers.indexOf('Nume');
  var colCurs1 = headers.indexOf('Acces Curs 1');
  var colCurs2 = headers.indexOf('Acces Curs 2');
  var colPachet = headers.indexOf('Acces Pachet');
  
  Logger.log('Indecsi coloane - Email: ' + colEmail + ', Parola: ' + colParola + ', Nume: ' + colNume);
  Logger.log('Acces Curs 1: ' + colCurs1 + ', Acces Curs 2: ' + colCurs2 + ', Acces Pachet: ' + colPachet);
  
  for (var i = 1; i < data.length; i++) {
    var emailSheet = data[i][colEmail];
    var parolaSheet = data[i][colParola];
    
    if (emailSheet === email && parolaSheet === parola) {
      Logger.log('[OK] LOGIN GASIT!');
      Logger.log('Nume: ' + data[i][colNume]);
      Logger.log('Acces Curs 1: ' + data[i][colCurs1]);
      Logger.log('Acces Curs 2: ' + data[i][colCurs2]);
      Logger.log('Acces Pachet: ' + data[i][colPachet]);
      return;
    }
  }
  
  Logger.log('[FAIL] LOGIN NU A FOST GASIT!');
}

function debugSession() {
  var sessionId = 'SESSION_ID_AICI';
  
  Logger.log('--- DEBUG SESSION: ' + sessionId + ' ---');
  
  try {
    var session1 = stripeApiRequest('checkout/sessions/' + sessionId, 'get', {});
    Logger.log('Session FARA expand:');
    Logger.log('Total Details: ' + JSON.stringify(session1.total_details, null, 2));
    
    var session2 = stripeApiRequest('checkout/sessions/' + sessionId, 'get', {
      expand: ['total_details.breakdown']
    });
    Logger.log('Session CU expand:');
    Logger.log('Total Details: ' + JSON.stringify(session2.total_details, null, 2));
    
    if (session2.total_details && session2.total_details.breakdown) {
      Logger.log('BREAKDOWN EXISTA!');
      Logger.log(JSON.stringify(session2.total_details.breakdown, null, 2));
      
      if (session2.total_details.breakdown.discounts) {
        for (var i = 0; i < session2.total_details.breakdown.discounts.length; i++) {
          Logger.log('Discount #' + (i + 1) + ': ' + JSON.stringify(session2.total_details.breakdown.discounts[i], null, 2));
        }
      }
    }
    
    if (session2.discounts) {
      Logger.log('session.discounts: ' + JSON.stringify(session2.discounts, null, 2));
    }
    
    Logger.log('[OK] DEBUG COMPLET!');
    
  } catch(error) {
    Logger.log('[EROARE] ' + error.toString());
  }
}

function diagnosticInstantAccess() {
  var testSessionId = 'cs_live_b12pOiBBdE2R2d0EUzB6MBn7mDe8wh1Le3eqIFD1uwVrKT1wJIeOjlvPAv'; 
  
  Logger.log('--- DIAGNOSTIC INSTANT ACCESS ---');
  
  try {
    var key = getStripeKey();
    Logger.log('[OK] Cheia Stripe configurata (incepe cu: ' + key.substring(0, 7) + '...)');

    Logger.log('Incerc procesarea pentru: ' + testSessionId);
    var rezultat = proceseazaSesiuneStripe(testSessionId);
    
    Logger.log('REZULTAT API: ' + JSON.stringify(rezultat, null, 2));

    if (rezultat.success) {
      Logger.log('[OK] Backend functioneaza corect.');
    } else {
      Logger.log('[FAIL] ' + rezultat.error);
    }

  } catch (e) {
    Logger.log('[EROARE CRITICA] ' + e.toString());
  }
}
