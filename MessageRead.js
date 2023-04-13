﻿const inputECO = document.getElementById("inpECO")
const btnCreateMsg = document.getElementById("btnCreateMsg")
const cbxConfig = document.getElementById("cbxConfig")
const sect1 = document.getElementById("searchDATE")

let ECOnumber = ""
let Configurator = ""

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        btnCreateMsg.onclick = closeConfig;
    }
});

async function closeConfig(){
    const mailItem = Office.context.mailbox.item;

    ECOnumber = inputECO.value
    Configurator = cbxConfig.value

    if (ECOnumber === "") {
        return null
    }
    if (Configurator === "") {
        return null
    }

    let CMailAddress = []

    switch (document.querySelector('#cbxConfig option:checked').parentElement.label) {
        case "WSHP & UV":
            CMailAddress=['Edgar.Huacuja@daikinapplied.com',
                            'David.Swift@daikinapplied.com',
                            'Francis.Middlemiss@daikinapplied.com',
                            'LEONARDO.RAMIREZ@daikinapplied.com']
            break;
        case "FC & BC":
            CMailAddress=['ANTONIO.AGUIRRE@daikinapplied.com',
                            'ALAN.CARRIZALES@daikinapplied.com',
                            'Tom.Simon@daikinapplied.com',
                            'Jorge.Zaragoza@daikinapplied.com',
                            'Oscar.Gutierrez@daikinapplied.com']
            break;
    }

    KBEETOTEAM = ['Javier.Macias@daikinapplied.com',
                    'EMMANUEL.HERRERA@daikinapplied.com',
                    'JOSE.CORNEJO@daikinapplied.com',
                    'ADAN.FERNANDEZ@daikinapplied.com',
                    'Toshiro.Fudizawa@daikinapplied.com',
                    'JOSE.TORRES@daikinapplied.com']

    mailItem.body.prependAsync(`<p style="font-family:'Arial';">Please do not release any order from the <b>${Configurator}</b> configurator <br> I will apply changes for ECO <b>${ECOnumber}</b>, Thanks!</p>`,{coercionType: Office.CoercionType.Html})
    mailItem.to.setAsync(CMailAddress)
    mailItem.cc.setAsync(KBEETOTEAM)
    mailItem.subject.setAsync(`RE: Freeze ETO for the ${Configurator} configurator. ECO ${ECOnumber}`)
}