const path = require('path')
const XLSX = require('xlsx')
const home = require('os').homedir()

let nbmoisactif = 3
let nbtranches = 5
let voirnoms = false
let entree = path.join(home, 'Desktop', 'pos.order.xls')
let sortie = path.join(home, 'Desktop', 'paniers.xlsx')

const coopid = 'Client/Nom affiché'
const datevente = 'Date de la commande'
const montant = 'Total'
const invit = '- INVITÉ(E)S,'

const coops = {}
const mois = {}
const lstmois = []
let numinvite = 1

try {
    for (let i = 0; i < process.argv.length; i++) {
        const a = process.argv[i]
        if (a.startsWith('entree=')) entree = a.substring('entree='.length)
        else if (a.startsWith('sortie=')) sortie = a.substring('sortie='.length)
        else if (a.startsWith('nbmoisactif=')) nbmoisactifs = parseInt(a.substring('nbmoisactif='.length), 10)
        else if (a.startsWith('nbtranches=')) nbtranches = parseInt(a.substring('nbtranches='.length), 10)
        else if (a == 'voirnoms=vrai') voirnoms = true
    }
    lecture()
    ecriture()
    process.exit(0)
} catch (e) {
    console.error(e.message)
    if (e.stack) console.error(e.stack)
    process.exit(1)
}

function ecriture() {
    const total = { mois:'9999', Mois:'9999', ventes: 0, nbA: 0, ventesC: 0, pmC: 0, nbI: 0, ventesI: 0, pmI: 0 }
    for (let i = 0; i < nbtranches; i++) total['p' + (i + 1)] = 0

    const resultat = []
    let lvm = []
    for (let i = 0; i < lstmois.length; i++) {
        const m = lstmois[i]
        const mo = mois[m]
        lvm = lvm.concat(mo.lvm)
        let r = { }
        r.mois = m
        r.Mois = m
        r.ventes = Math.round(mo.ventesC + mo.ventesI)
        r.nbA = mo.nbA
        r.ventesC = Math.round(mo.ventesC)
        total.ventesC += mo.ventesC
        r.pmC = Math.round(mo.ventesC ? mo.ventesC / mo.nbA : 0)
        r.nbI = mo.nbI
        r.ventesI = Math.round(mo.ventesI)
        r.pmI = Math.round(mo.ventesI ? mo.ventesI / mo.nbI : 0)
        total.ventesI += mo.ventesI
        for (let i = 0; i < nbtranches; i++) r['p' + (i + 1)] = mo.p[i]
        resultat.push(r)
        console.log(JSON.stringify(r))
    }
    const p = distrib(lvm)
    for (let i = 0; i < nbtranches; i++) total['p' + (i + 1)] = p[i]

    const lcoops = []
    let cpt = 1
    for (let c in coops) {
        const r = {}
        const coop = coops[c]
        r.coop = !voirnoms ? (coop.invite ? 'I' : 'C') + cpt++ : c
        for (let i = 0; i < lstmois.length; i++) {
            const m = lstmois[i]
            r[m] = coop[m] ? Math.round(coop[m]) : 0
        }
        if (coop.invite) {
            total.nbI++
        } else {
            total.nbA++
        }
        lcoops.push(r)
    }
    total.ventes = Math.round(total.ventesC + total.ventesI)
    total.pmC = Math.round(total.ventesC / total.nbA / lstmois.length)
    total.pmI = Math.round(total.ventesI / total.nbI)
    total.ventesC = Math.round(total.ventesC)
    total.ventesI = Math.round(total.ventesI)
    resultat.push(total)
    console.log(JSON.stringify(total))
    
    const workbook = XLSX.utils.book_new()
    const cols = ['mois', 'ventesC', 'ventesI', 'ventes', 'nbA', 'pmC', 'nbI', 'pmI', 'Mois']
    for (let i = 1; i <= nbtranches; i++) cols.push('p' + i)
    const ws = XLSX.utils.json_to_sheet(resultat, {header:cols})
    XLSX.utils.book_append_sheet(workbook, ws, "Paniers")
    const cols2 = ['coop']
    for (let i = 1; i <= lstmois; i++) cols2.push(lsmois[i])
    const ws2 = XLSX.utils.json_to_sheet(lcoops, {header:cols2})
    XLSX.utils.book_append_sheet(workbook, ws2, "Coops")
    XLSX.writeFile(workbook, sortie)
}

function nomCoop(n) {
    let i = n.indexOf('- ')
    const n1 = n.substring(i + 2)
    i = n1.indexOf(',')
    const j = n1.indexOf(',', i + 1)
    const n3 = j == -1 ? n1 : n1.substring(0, j)
    return n3
}

function lecture () {
    const wb = XLSX.readFile(entree, {cellDates:true})
    const ws = wb.Sheets[wb.SheetNames]
    const rows = XLSX.utils.sheet_to_json(ws, {blankrows:false})
    for (let i = 0; i < rows.length; i++) {
        const r = rows[i]
        const d = r[datevente]
        let c = r[coopid]

        const invite = c.indexOf(invit) !== -1
        if (invite) {
            c = '$$$' + numinvite++
        } else {
            c = nomCoop(c)
        }

        let coop = coops[c]
        if (!coop) {
            coop = { invite: invite }
            coops[c] = coop
        }

        const x = d.getMonth() + 1
        const m = '' + (d.getFullYear() % 100) + '-' + (x < 10 ? '0' + x : x)

        const v = r[montant]
        let mo = mois[m]
        if (mo === undefined) { 
            mo = { ventesC: 0, nbI:0, ventesI: 0, nbA: 0}
            mois[m] = mo
            lstmois.push(m)
        }        
        if (coop[m] === undefined) coop[m] = v; else coop[m] += v
    }

    lstmois.sort()

    for (let i = 0; i < lstmois.length; i++) {
        const av = []
        for (let j = i, k = nbmoisactif; j >= 0 && k > 0; j--, k--) 
            av.push(lstmois[j])
        mois[lstmois[i]].avant = av
    }

    for (let m in mois) {
        const im = lstmois.indexOf(m)
        const mo = mois[m]
        mo.lvm = []
        for (let c in coops) {
            const coop = coops[c]
            if (!coop.invite && estActif(coop, im)) mo.nbA++
            let v = coop[m]
            if (v !== undefined) {
                if (coop.invite) {
                    mo.nbI++
                    mo.ventesI += v
                } else {
                    mo.lvm.push(v)
                    mo.ventesC += v
                }
            }
        }
        let x = distrib(mo.lvm)
        mo.p = x
    }
}

function distrib(lvm) {
    lvm.sort((a,b) => { return a < b ? -1 : (a > b) ? 1 : 0})
    const p = new Array(nbtranches)
    p.fill(0)
    const tr = Math.floor(lvm.length / nbtranches)
    for (let i = 0, d = 0; i < nbtranches; i++) {
        const f = i == nbtranches - 1 ? lvm.length : d + tr
        let x = 0
        for (let j = d; j < f; j++) x += lvm[j]
        p[i] = Math.round(x / (f - d))
        d = f
    }
    return p
}

function estActif(coop, im) {
    const av = mois[lstmois[im]].avant
    for (let i = 0; i < av.length; i++) {
        const x = coop[av[i]]
        if (x) return true
    }
    return false
}
