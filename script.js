const inputECO = document.getElementById("inpECO")
const btnCreateMsg = document.getElementById("btnCreateMsg")
const cbxConfig = document.getElementById("cbxConfig")
const sect1 = document.getElementById("searchDATE")

let ECOnumber = ""
let Configurator = ""

function ShowHide() {
    if (document.getElementById("rbClose").checked === true) {
        inputECO.hidden = false
        document.getElementById("ttlECO").hidden = false
    } else {
        inputECO.hidden = true
        document.getElementById("ttlECO").hidden = true
    }
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        btnCreateMsg.onclick = closeConfig;
    }
});

async function closeConfig(){
    const mailItem = Office.context.mailbox.item;

    ECOnumber = inputECO.value
    Configurator = cbxConfig.value

    if (document.getElementById("rbClose").checked === true) {
        
    if (ECOnumber === "") {
        return null
    }

    let CMailAddress = []

    switch (document.querySelector('#cbxConfig option:checked').parentElement.label) {
        case "WSHP":
            CMailAddress=['Edgar.Huacuja@daikinapplied.com',
                            'DAVID.GALAVIZ@daikinmx.com',
                            'Cesar.Soto@daikinmx.com']
            break;
        case "UV":
            CMailAddress=['LEONARDO.RAMIREZ@daikinapplied.com',
                            'Martin.Ramos@daikinmx.com']
            break;
        case "FC":
            CMailAddress=['ANTONIO.AGUIRRE@daikinapplied.com',
                            'Jorge.Zaragoza@daikinapplied.com']
            break;
    }

    KBEETOTEAM = ['Javier.Macias@daikinapplied.com',
                    'EMMANUEL.HERRERA@daikinapplied.com',
                    'JOSE.CORNEJO@daikinapplied.com',
                    'ADAN.FERNANDEZ@daikinapplied.com',
                    'GLORIA.CASTANEDA@daikinmx.com',
                    'Angel.Gallegos@daikinmx.com']

    mailItem.body.prependAsync(`<p style="font-family:'Arial';">Please do not release any order from the <b>${Configurator}</b> configurator <br> I will apply changes for ECO <b>${ECOnumber}</b>, Thanks!</p>`,{coercionType: Office.CoercionType.Html})
    mailItem.to.setAsync(CMailAddress)
    mailItem.cc.setAsync(KBEETOTEAM)
    mailItem.subject.setAsync(`Freeze ETO for the ${Configurator} configurator. ECO ${ECOnumber}`)

    } 
    if(document.getElementById("rbOpen").checked === true){
        mailItem.body.prependAsync(`<p style="font-family:'Arial';"><b>${Configurator}</b> is ready to process orders again, Thank you!~</p>`,{coercionType: Office.CoercionType.Html})
    }

}