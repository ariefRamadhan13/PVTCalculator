import os
import numpy as np
import matplotlib.pyplot as plt
import openpyxl


def Hapus():
    os.system('cls')

def Garis():
    print('---------------------------------------------------')

###########################################################################################
###########################################################################################
def Tampilan():
    
    print('---------------------------------------------------------------\nSelamat datang diaplikasi perhitungan Properti Fluida Reservoir\n---------------------------------------------------------------\n1. Input data\n2. PVT Calculator\n3. Modifikasi Data\n4. Lihat data\n5. Keluar Aplikasi\n6. Lihat Grafik\n7. Simpan data ke Excel (Masuk PVT Calculator terlebih dahulu)')
    inputUser = input('---------------------------------------------------------------\nMasukkan Pilihan: ')
    Hapus()
    try:
        inputUser = int(inputUser)
        return inputUser
    except:
        print('Maaf anda salah input')
    Hapus()
    return inputUser

def TampilGrafik(x,y,Pb):
    plt.plot(x,y)
    plt.axvline(Pb,color='red',linestyle='dashed',linewidth=3)
    plt.show
    plt.pause(100)
    
    inputan = int(input('Apakah ingin lanjut ke aplikasi?\n1. Ya\n2. Tidak\nMasukkan pilihan: '))
    if inputan == 1:
        T = True
        Hapus()
        return T
    else:
        T = False
        print('Terimakasih telah menggunakan program kami')
        return T

###########################################################################################
def GeneralData():
    print('MASUKKAN GENERAL DATA YANG DIBUTUHKAN')
    Treservoir = float(input('Masukkan Temperatur Reservoir (Fahrenheit): ')) + 460
    Preservoir = float(input('Masukkan Tekanan Initial Reservoir (psia): '))
    Pstandar = float(input('Masukkan tekanan standar (psia): '))
    GasGravity = float(input('Masukkan Gas Gravity: '))
    Hapus()
    return Treservoir, Preservoir, Pstandar, GasGravity
###########################################################################################
def OilData(Tres,SgGas):
    print('MASUKKAN OIL DATA ')
    OilAPI = float(input('Masukkan API Oil: '))
    Pseparator = float(input('Masukkan tekanan separator (psia): '))
    Tseparator = float(input('Masukkan temperatur separator (Fahrenheit): '))
    Hapus()
    pilihanKorelasi = 1
    Hapus()
    print('Apa data yang tersedia?\n1. Pb\n2. Rs@Pb dengan Standing Correlation')
    pilihanUserOilData = int(input('Masukkan pilihan: '))
    Hapus()
    if pilihanUserOilData == 1:
        PbValue = float(input('Masukkan nilai Pb (psia): '))
        RsPbValue = 0
        Hapus()
        return OilAPI, Pseparator, Tseparator, PbValue, RsPbValue, pilihanKorelasi
    elif pilihanUserOilData == 2:#Metode Standing Correlation
        RsPbValue = float(input('Masukkan nilai Solution GOR|Pb (SCF/STB): '))/1000
        y = 0.00091*(Tres-460)-0.0125*(OilAPI)
        PbValue = 18*((RsPbValue*1000000/SgGas)**0.83)*pow(10,y)
        Hapus()
        return OilAPI, Pseparator, Tseparator, PbValue, RsPbValue, pilihanKorelasi
    Hapus()
###########################################################################################
def Impurities():
    print('MASUKKAN DATA % MOL MASING-MASING UNSUR')
    CO2 = float(input('Masukkan persentase fraksi mol CO2: '))
    N2 = float(input('Masukkan persentase fraksi mol N2: '))
    H2S = float(input('Masukkan persentase fraksi mol H2S: '))
    Hapus()
    return CO2/100, H2S/100, N2/100
###########################################################################################
def BrineData():
    print('Masukkan kondisi Brine\n1. Gas Saturated Brine\n2. Gas Free Brine')
    InputBrine = int(input('Masukkan pilihan: '))
    Hapus()
    TDS = float(input('Masukkan Persentase TDS: '))
    Hapus()
    return InputBrine, TDS
###########################################################################################
###########################################################################################

###########################################################################################
def Conditions(Pres, Pb):
    if Pres<Pb:
        print('Conditions = Saturated')
    elif Pres>Pb:
        print('Conditions = Undersaturated')
    elif Pres==Pb:
        print('Conditions = (Pb = Preservoir)')
###########################################################################################
def Bo(condSat, condUnsat, Rsg, SgGas, API, Tres, Pres, Pb, Rsb, Psep, Tsep, Bob):
    SG = 141.5 / (131.5 + API)
    Rsg = float(Rsg)
    Gammags = SgGas * (1 + 5.912 * pow(10, -5) * API * (Tsep) * np.log10(Psep / 114.7))
    if Pres <= Pb:
        Rs = Rsg * 1000
        if condSat == 1 or condSat == '1':# Standing
            BOil = 0.9759 + 0.00012 * pow(Rs * pow(SgGas / SG, 0.5) + 1.25 * (Tres - 460), 1.2)
            return BOil
        elif condSat == 2 or condSat == '2':#'Vasquez-Beggs
            if API > 30:
                C1 = 4.67 * pow(10, -4)
                C2 = 1.1 * pow(10, -5)
                C3 = 1.337 * pow(10, -9)
            else:
                C1 = 4.677 * pow(10, -4)
                C2 = 1.751 * pow(10, -5)
                C3 = -1.881 * pow(10, -8)
            BOil = 1 + C1 * Rs + (Tres - 520) * (API / Gammags) * (C2 + C3 * Rs)
            return BOil
        elif condSat == 3 or condSat == '3':#'Glaso
            Bobb = Rs * pow(SgGas / SG, 0.526) + 0.968 * (Tres - 460)
            A = -6.58511 + 2.91329 * np.log10(Bobb) - 0.27683 * pow(np.log10(Bobb), 2)
            BOil = 1 + pow(10, A)
            return BOil
    else:
        Rs = Rsg * 1000
        Rsb = Rsb * 1000
        if condUnsat == 1 or condUnsat == '1': #'Vasquez-Beggs
            A1 = pow(10, -5) * (-1433 + 5 * Rsb + 17.2 * (Tres - 460) - 1180 * Gammags + 12.61 * API)
            BOil = Bob * np.exp(-A1 * np.log(Pres / Pb))
            return BOil
        elif condUnsat == 2 or condUnsat == '2':#'Petrosky-Farshad
            A1 = 4.1646 * pow(10, -7) * pow(Rsb, 0.69357) * pow(SgGas, 0.1885) * pow(API, 0.3272) * pow(Tres - 460, 0.6729)
            BOil = Bob * np.exp(-A1 * (pow(Pres, 0.4094) - pow(Pb, 0.4096)))
            return BOil
###########################################################################################
def Rs(cond,Pres,SgGas,T,API,Psep,Tsep,Pb):
    if cond == 1:#'Standing'
        P = min(Pres, Pb)
        X = 0.0125 * API - 0.00091 * (T - 460)
        RsOil = SgGas * pow((P / 18.2 + 1.4) * pow(10, X), 1.2048) / 1000
        return RsOil
    elif cond == 2:#'Vazques & Beggs
        P = min(Pres, Pb)
        SG = 141.5 / (API + 131.5)
        Gammags = SgGas * (1 + 5.912 * pow(10, -5) * API * (Tsep) * np.log10(Psep / 114.7))
        C1 = (0, 0.0362, 0.0178)
        C2 = (0, 1.0937, 1.187)
        C3 = (0, 25.724, 23.931)
        if API <= 30:
            RsOil = C1[1] * Gammags * (P**C2[1]) * np.exp(C3[1] * API / T) / 1000
            return RsOil
        else:
            RsOil = C1[2] * Gammags * (P**C2[2]) * np.exp(C3[2] * API / T) / 1000
            return RsOil

    elif cond == 3:#'Glaso
        P = min(Pres, Pb)
        Pbst = pow(10, 2.8869 - pow(14.1811 - 3.3093 * np.log10(P), 0.5))
        RsOil = SgGas / 1000 * pow(pow(API, 0.989) * Pbst / pow(T - 460, 0.172), 1.2255)
        return RsOil
###########################################################################################
def OilDens(condSat, condUsat , API , Rs , Rsb , Bo , Co , SgGas , Pres , Pb , Bob , Tres , Tsep , Psep , ODB):
    SG = 141.5 / (API + 131.5)
    Gammags = SgGas * (1 + 5.912 * pow(10, -5) * API * (Tsep - 460) * np.log10(Psep / 114.7))
    if Pres <= Pb:
        if condSat == 1 or condSat == '1':#'General
            OilDens = (62.4 * SG + 0.0136 * Rs * SgGas) / Bo
            return OilDens
        elif condSat == 2 or condSat == '2':#'Standing
            OilDens = (62.4 * SG + 0.0136 * Rs * SgGas) / (0.972 + 0.000147 * pow(Rs * pow(SgGas / SG, 0.5) + 1.25 * (Tres - 460), 1.175))
            return OilDens
    else:
        if condUsat == 1 or condUsat =='1': #'Vasquez
            A = pow(10, -5) * (-1433 + 5 * Rs + 17.2 * (Tres - 460) - 1180 * Gammags + 12.61 * API)
            OilDens = ODB * np.exp(A * np.log10(Pres / Pb))
            return OilDens
        elif condUsat == 2 or condUsat =='2':#'Ahmed
            A = -0.00018473
            B = -1 / (4.588893 + 0.0025999 * Rs)
            OilDens = ODB * np.exp(B * (np.exp(A * Pres) - np.exp(A * Pb)))
            return OilDens
###########################################################################################
def mu0(cond,mu0d,Rsinput,Pres,Pb,OilVisPb):
    Rs = Rsinput * 1000
    if Pres <= Pb:
        if cond == 1 or cond == '1':#'Chew-Connally
            E = 3.74 * pow(10, -3) * Rs
            D = 1.1 * pow(10, -3) * Rs
            C = 8.62 * pow(10, -5) * Rs
            B = 0.68 / pow(10, C) + 0.25 / pow(10, D) + 0.062 / pow(10, E)
            A = Rs * (2.2 * pow(10, -7) * Rs - 7.4 * pow(10, -4))
            VisOil = pow(10, A) * pow(mu0d, B)
            return VisOil
        elif cond == 2 or cond == '2': #'Beggs-Robinson
            A = 10.715 * pow(Rs + 100, -0.515)
            B = 5.44 * pow(Rs + 150, -0.338)
            VisOil = A * pow(mu0d, B)
            return VisOil
    elif Pres > Pb:# 'Vazquez Beggs Under Saturated
        A = -3.9 * pow(10, -5) * Pres - 5
        m = 2.6 * pow(Pres, 1.187) * pow(10, A)
        VisOil = OilVisPb * pow(Pres/Pb, m) *pow(10,-2)
        return VisOil
###########################################################################################
def mu0d(Cond, API, T):
    if Cond == 1 or Cond == '1':#Beals
        A = pow(10, 0.43 + 8.33 / API)
        DeadVisOil = (0.32 + 1.8 * 10000000 / pow(API, 4.53)) * pow(360 / (T - 260), A)
        return DeadVisOil
    elif Cond == 2 or Cond == '2':#Glaso
        A = 10.313 * np.log10(T - 460) - 36.447
        DeadVisOil = 3.141 * pow(10, 10) * pow(T - 460, -3.444) * pow(np.log10(API), A)
        return DeadVisOil
    elif Cond == 3 or Cond == '3':#Beggs-Robinson
        Z = 3.0324 - 0.02023 * API
        Y = pow(10, Z)
        X = Y * pow(T - 460, -1.163)
        DeadVisOil = pow(10, X) - 1
        return DeadVisOil
###########################################################################################
def Co(condSat, condUsat , Rs , Rsb, Pres, Tres, API , SgGas, Pb, Bg, Bo, Psep, Tsep):
    SG = 141.5 / (API + 131.5)
    Gammags = SgGas * (1 + 5.912 * pow(10, -5) * API * (Tsep - 460) * np.log10(Psep / 114.7))
    if Pres <= Pb:
        # SATURATED
        if condSat == 1 or condSat =='1': #'McCain
            A = -7.573 - 0.383 * np.log10(Pb) - 1.45 * np.log10(Pres) + 1.402 * np.log10(Tres-460) + 0.256 * np.log10(API) + 0.449 * np.log10(Rsb)
            IThermCompress = np.exp(A)
            return IThermCompress*100000
        elif condSat == 2 or condSat =='2': # 'Standing
            IThermCompress = -Rs / (Bo * (0.83 * Pres + 21.75)) * (0.00014 * np.sqrt(SgGas / SG) * pow(Rs * np.sqrt(SgGas / SG) + 1.25 * (Tres - 460), 0.12) - Bg)
            return IThermCompress*100000
    else:
        # UNDERSATURATED
        if condUsat == 1 or condUsat =='1': #'Vasquez-Beggs
            IThermCompress = (-1433 + 5 * Rsb + 17.2 * (Tres - 460) - 1180 * Gammags + 12.61 * API) / (pow(10, 5) * Pres)
            return IThermCompress*100000
        elif condUsat == 2 or condUsat =='2': #'PetroskyFarshad
            IThermCompress = 1.705 * pow(10, -7) * pow(Rsb, 0.69357) * pow(SgGas, 0.1885) * pow(API, 0.3272) * pow(Tres - 460, 0.6729) * pow(1 / Pres, 0.5906)
            return IThermCompress*100000
###########################################################################################
###########################################################################################

def Ppc(x,y):
    if x == 1:# Natural Gasses - Standing
        Ppc = 677 + 15*y - 37.5 * (y**2)
        return Ppc
    elif x == 2:#Condensate Gases - Sutton
        Ppc = 756.8 - 131 * y - 3.6 * (y**2)
        return Ppc
    elif x == 3:#Condensate Gases - Standing
        Ppc = 706 - 51.7 * y - 11.1 * (y**2)
        return Ppc
###########################################################################################
def Tpc(x,y):
    if x == 1:# Natural Gasses - Standing
        Ppc = 168 + 325*y - 12.5 * (y**2)
        return Ppc
    elif x == 2:#Condensate Gases - Sutton
        Ppc = 169.2 + 349.5 * y - 74 * (y**2)
        return Ppc
    elif x == 3:#Condensate Gases - Standing
        Ppc = 187 + 330 * y - 71.5 * (y**2)
        return Ppc
###########################################################################################
def TpcCorrection(Cond,TP,Tpc,Ppc,H2S,CO2,N2):
    if Cond == 1:#Wichert-Aziz Correction Method
        B = H2S
        A = H2S + CO2
        eps = 120 * (pow(A, 0.9) - pow(A, 1.6)) + 15 * (pow(B, 0.5) - pow(B, 4))
        Tcorrect = Tpc - eps
        Pcorrect = (Ppc * Tcorrect) / (Tpc + B * (1 - B) * eps)
        if TP == 1:
            return Tcorrect
        elif TP == 2:
            return Pcorrect
    elif Cond == 2:#Carr-Kobayashi-Burrows(1954) Method
        Tcorrect = Tpc - 80 * CO2 + 130 * H2S - 250 * N2
        Pcorrect = Ppc + 440 * CO2 + 600 * H2S - 170 * N2
        if TP == 1:
            T = Tcorrect
            return T
        elif TP == 2:
            P = Pcorrect
            return P
########################################################################################### 
###########################################################################################     
def Y_HY(Y, Tpr, P):
    T = 1 / Tpr
    bag1 = -0.06125 * P * T * 2.71828 ** (-1.2 * (1 - T) ** 2)
    bag2 = (Y + Y ** 2 + Y ** 3 - Y ** 4) / (1 - Y) ** 3
    bag3 = (14.76 * T - 9.76 * T * T + 4.58 * T * T * T) * Y ** 2
    bag4 = (90.7 * T - 242.2 * T * T + 42.4 * T * T * T) * (Y ** (2.18 + 2.82 * T))
    
    Y_HY = (bag1 + bag2 - bag3 + bag4)
    return Y_HY
# #########################################################################################    
def ZFactor(Pilihan,Tpr, Ppr):
    if Pilihan == 1 or Pilihan == '1': # Dranchuk-Abu-Kassem (1975)
        A1 = 0.3265
        A2 = -1.07
        A3 = -0.5339
        A4 = 0.01569
        A5 = -0.05165
        A6 = 0.5475
        A7 = -0.7361
        A8 = 0.1844
        A9 = 0.1056
        A10 = 0.6134
        A11 = 0.721

        R1 = A1 + A2 / Tpr + A3 / pow(Tpr, 3) + A4 / pow(Tpr, 4) + A5 / pow(Tpr, 5)
        R2 = 0.27 * Ppr / Tpr
        R3 = A6 + A7 / Tpr + A8 / pow(Tpr, 2)
        R4 = A9 * (A7 / Tpr + A8 / pow(Tpr, 2))
        R5 = A10 / pow(Tpr, 3)

        Rho = 0.27 * Ppr / Tpr
        Rhobef = Rho
        for i in range (1,150):
            frho = R1 * Rho - R2 / Rho + R3 * pow(Rho, 2) - R4 * pow(Rho, 5) + R5 * (1 + A11 * pow(Rho, 2)) * pow(Rho, 2) * np.exp(-A11 * pow(Rho, 2)) + 1
            dfrho = R1 + R2 / pow(Rho, 2) + 2 * R3 * Rho - 5 * R4 * pow(Rho, 4) + 2 * R5 * Rho * np.exp(-A11 * pow(Rho, 2)) * ((1 + 2 * A11 * pow(Rho, 3)) - A11 * pow(Rho, 2) * (1 + A11 * pow(Rho, 2)))
            Rho = Rho - frho / dfrho
            test = np.abs((Rho - Rhobef) / Rho)
            if test < 0.00001:
                Rhobef = Rho
        Z_DAK = 0.27 * Ppr/(Rho*Tpr)
        return Z_DAK
    elif Pilihan == 2 or Pilihan == '2': # Dranchuk-Purvis-Robinson (1974)
        A = 0.064225133
        B = 0.53530771 * Tpr - 0.61232032
        C = 0.31506237 * Tpr - 1.0467099 - 0.57832729 / Tpr**2
        D = Tpr
        E = 0.68157001 / Tpr**2
        F = 0.68446549
        G = 0.27 * Ppr

        Rho = 0.27 * Ppr / Tpr  # Initial guess
        Rhoold = Rho
        for i in range (1,100):
            frho = A * Rho**6 + B * Rho**3 + C * Rho**2 + D * Rho + E * Rho**3 * (1 + F * Rho**2) * np.exp(-F * Rho**2) - G
            dfrho = 6 * A * Rho**5 + 3 * B * Rho**2 + 2 * C * Rho + D + E * Rho**2 * (3 + F * Rho**2 * (3 - 2 * F * Rho**2)) * np.exp(-F * Rho**2)
            Rho = Rho - frho / dfrho
            test = np.abs((Rho - Rhoold) / Rho)
            if test < 0.00001:
                Rhobef = Rho
        Z_DPR = 0.27 * Ppr/(Rho*Tpr)
        return Z_DPR
    
    elif Pilihan == 3:#'Z by Hall-Yarborough
        Yo = 0.01
        Ybef = Yo - 1
        MaxIter = 125
        i = 0
        while abs(Ybef - Yo) > 0.00001 and i <= MaxIter:
            Ybef = Yo
            Yo = Yo - (Y_HY(Yo, Tpr, Ppr) * 0.00001 / (Y_HY((Yo + 0.00001), Tpr, Ppr) - Y_HY(Yo, Tpr, Ppr)))
            i = i + 1
        Y = Yo
        Zmenu = (0.06125 * Ppr / (Tpr * Y)) * np.exp(-1.2 * pow((1 - 1 / Tpr), 2))
        return Zmenu
    
    elif Pilihan == 4 or Pilihan == '4': #Brill and Beggs
        A = 1.39 * pow(Tpr - 0.92, 0.5) - 0.36 * Tpr - 0.1
        B = (0.62 - 0.23 * Tpr) * Ppr + (0.066 / (Tpr - 0.86) - 0.037) * pow(Ppr, 2) + 0.32 / pow(10, 9 * (Tpr - 1)) * pow(Ppr, 6)
        C = 0.132 - 0.32 * np.log10(Tpr)
        D = pow(10, 0.3106 - 0.49 * Tpr + 0.1824 * pow(Tpr, 2))
        
        Z_BB = A + (1 - A) / np.exp(B) + C * pow(Ppr, D)
        return Z_BB
###########################################################################################
def Zmenu(cond, T, Tpc,P, Ppc):
    Tpr = T / Tpc
    Ppr = P / Ppc
    if cond == 1 or cond == '1':#'Z by Dranchuk-Abou-Kassem
        Zmenu = ZFactor(cond,Tpr, Ppr)
        return Zmenu
    elif cond == 2 or cond == '2':#'Z by Dranchuk-Purvis-Robinson
        Zmenu = ZFactor(cond, Tpr, Ppr)
        return Zmenu
    elif cond == 3:#'Z by Hall-Yarborough
        Yo = 0.01
        Ybef = Yo - 1
        MaxIter = 125
        i = 0
        while abs(Ybef - Yo) > 0.00000001 and i <= MaxIter:
            Ybef = Yo
            Yo = Yo - (Y_HY(Yo, Tpr, Ppr) * 0.00001 / (Y_HY((Yo + 0.00001), Tpr, Ppr) - Y_HY(Yo, Tpr, Ppr)))
            i = i + 1
        Y = Yo
        Zmenu = (0.06125 * Ppr / (Tpr * Y)) * np.exp(-1.2 * pow((1 - 1 / Tpr), 2))
        return Zmenu
    
    elif cond == 4:#'Z by Beggs-Brill
        Zmenu = ZFactor(cond, Tpr, Ppr)
        return Zmenu
###########################################################################################
def Bg(Z,T,P):
    Bg = 0.005035 * (Z * T) * 1000 / P
    return Bg
###########################################################################################
def H20inGas(S,T,P):
    PvW = np.exp(69.103501 - 13064.76 / T - 7.3037 * np.log(T) + 1.2856 * pow(10, -6) * pow(T, 2))
    mw = 18 / 380 * PvW / P
    #'No correction for MW
    corr = 1 - 4.92 * pow(10, -3) * S - 1.7672 * pow(10, -4) * pow(S, 2)
    WatContent = mw * corr
    return WatContent*1000
###########################################################################################
def GasViscosity(cond,Tpc, Ppc,yCO2,yN2,yH2S,SgGas,T,P,RhoG):
    Tpr = T / Tpc
    Ppr = P / Ppc
    Ma = SgGas * 28.96

    if cond == 1:#Carr-Kobayashi Viscosity of Gas By Standing
        mu1nc = (1.709 * (pow(10, -5) - 2.062 * pow(10, -6) * SgGas) * (T - 460) + 8.1118 * pow(10, -3) - 6.15 * pow(10, -3) * np.log10(SgGas))
        muc = yCO2 * ((9.08 * pow(10, -3)) * np.log10(SgGas) + (6.24 * pow(10, -3)))
        mun = yN2 * (8.48 * pow(10, -3) * np.log10(SgGas) + 9.59 * pow(10, -3))
        mus = yH2S * (8.49 * pow(10, -3) * np.log10(SgGas) + 3.73 * pow(10, -3))
        mu1 = mu1nc + muc + mun + mus
        #Viscosity of Gas By Dempsey
        A0 = -2.4621182
        A1 = 2.970547414
        A2 = -2.86264054 * (10 ** -1)
        A3 = 8.05420522 * (10 ** -3)
        A4 = 2.80860949
        A5 = -3.49803305
        A6 = 3.6037302 * (10 ** -1)
        A7 = -1.044324 * (10 **-2)
        A8 = -7.93385648 * (10 ** -1)
        A9 = 1.39643306
        A10 = -1.49144925 * (10 ** -1)
        A11 = 4.41015512 * (10 **-3)
        a12 = 8.39387178 * (10 ** -2)
        a13 = -1.86408848 * (10 ** -1)
        a14 = 2.03367881 * (10 ** -2)
        A15 = -6.09579263 * (10 ** -4)
    
        Logn1 = A0 + A1 * Ppr + A2 * Ppr**2 + A3 * Ppr**3
        Logn2 = Tpr * (A4 + A5 * Ppr + A6 * Ppr**2 + A7 * Ppr**3)
        Logn3 = pow(Tpr, 2) * (A8 + A9 * Ppr + A10 * Ppr**2 + A11 * Ppr**3)
        Logn4 = pow(Tpr, 3) * (a12 + a13 * Ppr + a14 * Ppr**2 + A15 * Ppr**3)
        
        Logn = Logn1 + Logn2 + Logn3 + Logn4
        
        VisGas = np.exp(Logn) * mu1 / Tpr
        return VisGas
        
    elif cond ==2:#By  Lee-Gonzalez-Eakin
        K = (9.4 + 0.02 * Ma) * pow(T, 1.5) / (209 + 19 * Ma + T)
        X = 3.5 + 986 / T + 0.01 * Ma
        Y = 2.4 - 0.2 * X
        
        VisGas = pow(10, -4) * K * np.exp(X * pow(RhoG/62.4, Y))
        return VisGas
###########################################################################################
def GasDensity(SgGas, T, Pres, Z):
    Ma = SgGas * 28.96
    GasDens = Pres * Ma / (Z * 10.73 * T)
    return GasDens
###########################################################################################
def Cg(cond,T,Tpc,P,Ppc):
    fz = (Zmenu(cond, T, Tpc, P + 0.0001, Ppc) - Zmenu(cond, T, Tpc, P, Ppc)) / 0.0001
    Cg = (1 / P) - (fz / Zmenu(cond, T, Tpc, P, Ppc))
    return Cg*100000
############################################################################################

###########################################################################################
###########################################################################################
###########################################################################################
def Bw(Cond, T, P):
    if Cond == 2 or Cond == '2': #Gas free brine
        A1 = 0.9947 + 0.58 * pow(10, -6) * (T - 460) + 1.02 * pow(10, -6) * (T - 460)**2
        A2 = -4.228 * pow(10, -6) + 1.8376 * pow(10, -8) * (T - 460) - 6.77 * pow(10, -11) * (T - 460)**2
        A3 = 1.3 * pow(10, -10) - 1.3855 * pow(10, -12) * (T - 460) + 4.285 * pow(10, -15) * (T - 460)**2
    elif Cond == 1 or Cond == '1': # Gas saturated brine
        A1 = 0.9911 + 6.35 * pow(10, -5) * (T - 460) + 8.5 * pow(10, -7) * (T - 460)**2
        A2 = -1.093 * pow(10, -6) - 3.497 * pow(10, -9) * (T - 460) + 4.57 * pow(10, -12) * (T - 460)**2
        A3 = -5 * pow(10, -11) + 6.429 * pow(10, -13) * (T - 460) - 1.43 * pow(10, -15) * (T - 460)**2
    Bw = A1 + A2 * P + A3 * P**2
    return Bw
###########################################################################################
def Rsw(Tres, P, TDS):
    T = Tres - 460
    A0 = 8.15839 - 6.12265 * pow(10, -2) * T + 1.91663 * pow(10, -4) * pow(T, 2) - 2.1654 * pow(10, -7) * pow(T, 3)
    A1 = 1.01021 * pow(10, -2) - 7.44241 * pow(10, -5) * T + 3.05553 * pow(10, -7) * pow(T, 2) - 2.94883 * pow(10, -10) * pow(T, 3)
    A2 = -pow(10, -7) * (9.02505 - 0.130237 * T + 8.53425 * pow(10, -4) * pow(T, 2) - 2.34122 * pow(10, -6) * pow(T, 3) + 2.37049 * pow(10, -9) * pow(T, 4))
    Rw = A0 + A1 * P + A2 * pow(P, 2)
    corr = pow(10, -0.0840655 * TDS * pow(T, -0.285854))
    RSwat = Rw * corr
    return RSwat/1000
###########################################################################################
def BrineDensity(TDS,Bw):
    DensWStd = 62.368 + 0.438603 * TDS + 1.60074 * pow(10, -3) * pow(TDS, 2)
    DensW = DensWStd / Bw
    return DensW
###########################################################################################
def BrineViscosity(cond,S,Pres,T):
    if cond == 1 or cond == '1': #Meehan
        D = 1.12166 - 0.0263951 * S + 6.79461 * pow(10, -4) * pow(S, 2) + 5.47119 * pow(10, -5) * pow(S, 3) - 1.55586 * pow(10, -6) * pow(S, 4)
        VisWT = (109.574 - 8.40564 * S + 0.313314 * pow(S, 2) + 8.72213 * pow(10, -3) * pow(S, 3)) * pow(T - 460, -D)
        VisW = VisWT * (0.9994 + 4.0295 * pow(10, -5) * Pres + 3.1062 * pow(10, -9) * pow(Pres, 2))
        return VisW
    elif cond == 2 or cond == '2': #Beggs-Brill
        VisW = np.exp(1.003 - 1.479 * pow(10, -2) * (T - 460) + 1.982 * pow(10, -5) * pow(T - 460, 2))
        return VisW
    
###########################################################################################
def CWater(Cond, Tres, P,Rsw,Tds):
    T = Tres - 460
    C0 = 3.8546 - 0.000134 * P
    C1 = -0.01052 + 0.000000477 * P
    C2 = 0.000039267 - 0.00000000088 * P
    Cwf = (C0 + C1 * T + C2 * pow(T, 2)) / 1000000
    corrSalt = (-0.052 + 0.00027 * T - 0.00000114 * pow(T, 2) + 0.000000001121 * pow(T, 3)) * Tds + 1
    if Cond == 2 or Cond == '2': #Free Brine
        CWater = Cwf * corrSalt * 100000
        return CWater
    elif Cond == 1 or Cond == '1': #Saturated Brine
        corrGas = (1 + 0.0089 * Rsw)
        CWater = Cwf * corrSalt * corrGas * 100000
        return CWater
###########################################################################################
def LihatData(a,b,c,d,e,f,g,h):
    print('DATA GENERAL')
    print(f'a. Temperatur Reservoir : {a[0]-460} Fahrenheit ({a[0]} Rankine)\nb. Tekanan Initial Reservoir (psia) : {a[1]}\nc. Tekanan standar (psia) : {a[2]}\nd. Gas Gravity : {a[3]}')
    Garis()
    print(f'OIL DATA\na. API Oil : {b[0]}\nb. Tekanan Separator (psia) : {b[1]}\nc. Temperatur Separator (Fahrenheit) : {b[2]}\nd. Pb Value : {b[3]}\ne. Data yang tersedia : {b[5]} (1. Pb| 2. Rs@Pb dengan Standing Correlation')
    Garis()
    print(f'DATA IMPURITIES\na. Fraksi CO2 : {c[0]*100}%\nb. Fraksi H2S : {c[1]*100}%\nc. Fraksi N2 : {c[2]*100}%')
    Garis()
    print(f'KONDISI BRINE\na. Persentase TDS : {d[1]}%\nb. Kondisi Brine : {d[0]} (1. Gas Saturated Brine | 2. Gas Free Brine)')
    Garis()
    print(f'KORELASI-KORELASI\na. Correlation Method untuk penentuan Tpc dan Ppc : {e} (1. Wichert-Aziz|2. Carr-Kobayashi-Burrow)')
    print(f'b. Correlation Method untuk penentuan Tpc dan Ppc Correction : {f} (1. Natural Gasses - Standing|2. Condensate Gases - Sutton|3. Condensate Gases - Standing)')
    print(f'c. Correlation Method untuk penentuan Z Factor : {g} (1. Dranchuk-Abu-Kassem|2. Dranchuk-Purvis-Robinson|3. Hall-Yarborough|4. Beggs-Brill)')
    print(f'd. Correlation Method untuk penentuan Gas Viscosity : {h} (1. Carr-Kobayashi-Burrow|2. Lee Gonzales Eakin)')
###########################################################################################
#                                       KORELASI
###########################################################################################
# A. Korelasi Oil Properties
# Undersaturated
def KorelasiUndersaturated():
    print('INPUT KORELASI PADA KONDISI UNDERSATURATED')
    CorBo = int(input('Masukkan korelasi untuk Bo\n1. Vasquez-Beggs\n2. Petrosky-Farshad \nMasukkan pilihan korelasi: '))
    Hapus()
    print('INPUT KORELASI PADA KONDISI UNDERSATURATED')
    CorRho = int(input('Masukkan korelasi untuk Rho0\n1. Vasquez-Beggs\n2. Ahmed\nMasukkan pilihan korelasi: '))
    Hapus()
    print('INPUT KORELASI PADA KONDISI UNDERSATURATED')
    CorCo = int(input('Masukkan korelasi untuk Co\n1. Vasquez-Beggs\n2. Petrosky-Farshad\nMasukkan pilihan korelasi: '))
    Hapus()
    return CorBo, CorRho, CorCo
# Saturated
def KorelasiSaturated():
    print('INPUT KORELASI PADA KONDISI SATURATED')
    CorBo = int(input('Masukkan korelasi untuk Bo\n1. Standing\n2. Vasquez-Beggs\n3. Glaso\nMasukkan pilihan korelasi: '))
    Hapus()
    print('INPUT KORELASI PADA KONDISI SATURATED')
    CorRs = int(input('Masukkan korelasi untuk Rs\n1. Standing\n2. Vasquez-Beggs\n3. Glaso\nMasukkan pilihan korelasi: '))
    Hapus()
    print('INPUT KORELASI PADA KONDISI SATURATED')
    CorRho = int(input('Masukkan korelasi untuk Rho0\n1. General\n2. Standing\nMasukkan pilihan korelasi: '))
    Hapus()
    print('INPUT KORELASI PADA KONDISI SATURATED')
    CorMu0 = int(input('Masukkan korelasi untuk mu0\n1. Chew Connally\n2. Beggs-Robinson\nMasukkan pilihan korelasi: '))
    Hapus()
    print('INPUT KORELASI PADA KONDISI SATURATED')
    CorMu0d = int(input('Masukkan korelasi untuk mu0d\n1. Beals\n2. Glaso\n3. Beggs-Robinson\nMasukkan pilihan korelasi: '))
    Hapus()
    print('INPUT KORELASI PADA KONDISI SATURATED')
    CorCo = int(input('Masukkan korelasi untuk Co\n1. McCain\n2. Standing\nMasukkan pilihan korelasi: '))
    Hapus()
    return CorBo, CorRs, CorRho, CorMu0, CorMu0d, CorCo
# B. Korelasi untuk Critical Properties
def CorrectionCriticalProperties():
    print("Pilih korelasi untuk Tpc' dan Ppc'\n1. Wichert-Aziz\n2. Carr-Kobayashi-Burrow")
    X = int(input('Masukkan pilihan: '))
    Hapus()
    return X

def TPCPPCCorrelation():
    print('Pilih Correlation Method untuk penentuan Tpc dan Ppc\n1. Natural Gasses - Standing\n2. Condensate Gases - Sutton\n3. Condensate Gases - Standing')
    CorelationTPC = int(input('Masukkan Pilihan: '))
    Hapus()
    return CorelationTPC

# C. Korelasi untuk Gas Properties
def ZCorrelation():
    print('Pilih Correlation Method untuk penentuan Z Factor\n1. Dranchuk-Abu-Kassem\n2. Dranchuk-Purvis-Robinson\n3. Hall-Yarborough\n4. Beggs-Brill')
    Zcor = int(input('Masukkan Pilihan: '))
    Hapus()
    return Zcor

def GasViscosityCorrelation():
    print('Pilih Correlation Method untuk penentuan Gas Viscosity\n1. Carr-Kobayashi-Burrow\n2. Lee Gonzales Eakin')
    GasViscoscor = int(input('Masukkan Pilihan: '))
    Hapus()
    return GasViscoscor

# D. Korelasi untuk Brine Properties
def BrineViscosityCorrelation():
    print('Pilih Correlation Method untuk penentuan Brine Viscosity\n1. Meehan\n2. Beggs-Brill')
    GasViscoscor = int(input('Masukkan Pilihan: '))
    Hapus()
    return GasViscoscor

###########################################################################################
###########################################################################################
###########################################################################################
T = True
# Simpan data Crude oil
Preservoir = 0
DataGeneral = ()
DataOil = ()
DataImpurities = ()
DataBrine = ()
CorrelationTPCPPC = ()
CorrectionCritProp = ()
ZCor = ()
GasVisCor = ()
DataKorelasiSaturated = (0,0,0,0,0,0)
DataKorelasiUndersaturated = (0,0,0)
DataBrineViscosityCorrelation = ()
HasilBg = 0
HasilRsPb = 0
Hasilmu0d = 0
HasilVisOilPb = 0
HasilBoPb = 0
Hapus()

###########################################################################################
###########################################################################################
# TEMPAT NYIMPAN TABEL DALAM LIST
PTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
SatCondTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
BoTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,14.7]
RsTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,14.7]
BwTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
RswTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
H20GasTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
BrineDensityTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
BrineViscosityTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
BgTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
EgTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
ZTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
GasDensTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
CwTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
OilViscosityTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
CgTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
GasVisTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
CoTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

def P(Pinitial,Pb,Pstandar):    
    P = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    P[0] = round(Pinitial,3)
    P[12] = round(Pb,3)
    P[11] = round(P[12]+1,3)
    P[13] = round(P[12]-1,3)
    P[30] = int(Pstandar)
    for i in range (1,11):
        P[i] = round(P[i-1]+((P[11]-P[0])/11),3)
    for i in range (14,30):
        P[i] = round(P[i-1]+((P[30]-P[13])/17),3)
    return P

def SatCond(P,Pb):
    SatCondTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        if P[i]<Pb:
            SatCondTabel[i] = 'Saturated'
        else:
            SatCondTabel[i] = 'Undersaturated'
    return SatCondTabel

def BoPb(condSat,condUnsat,Rsg,SgGas,API,Tres,Pres,Pb,RsPb,Psep,Tsep,BopbFalse):
    BoTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    BoPb = Bo(condSat,condUnsat,Rsg[12],SgGas,API,Tres,Pres[12],Pb,RsPb,Psep,Tsep,0)
    BoTabel[12] = Bo(condSat,condUnsat,Rsg[12],SgGas,API,Tres,Pres[12],Pb,RsPb,Psep,Tsep,0)
    for i in range(12):
        BoTabel[i] = Bo(condSat,condUnsat,Rsg[i],SgGas,API,Tres,Pres[i],Pb,RsPb,Psep,Tsep,BoPb)
    for i in range(13,31):
        BoTabel[i] = Bo(condSat,condUnsat,Rsg[i],SgGas,API,Tres,Pres[i],Pb,RsPb,Psep,Tsep,BoPb)
    return BoTabel


def Rsb(cond,PTabel,SgGas,T,API,Psep,Tsep,Pb):
    RsTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range (31):
        Pembulatan = Rs(cond,PTabel[i],SgGas,T,API,Psep,Tsep,Pb)
        RsTabel[i] = Pembulatan
    return RsTabel

def BwPb(cond,Tres,PTabel):
    BwTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        BwTabel[i] = Bw(cond,Tres,PTabel[i])
    return BwTabel

def RswPb(Tres,PTabel,TDS):
    RswTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        RswTabel[i] = Rsw(Tres,PTabel[i],TDS)
    return RswTabel

def H20Gas(S,T,P):
    H20GasTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        H20GasTabel[i] = H20inGas(S,T,P[i])
    return H20GasTabel

def BrineDensityT(S,BwT):
    BrineDensityTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        BrineDensityTabel[i] = BrineDensity(S,BwT[i])
    return BrineDensityTabel

def BrineViscosityT(cond,S,PTabel,Tres):
    TDS = S/100
    BrineViscosityTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        BrineViscosityTabel[i] = BrineViscosity(cond,TDS,PTabel[i],Tres)
    return BrineViscosityTabel

def BgT(Z,T,P):
    BgTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        BgTabel[i] = Bg(Z[i],T,P[i])
    return BgTabel

def EgT(BgPb):
    EgTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        EgTabel[i] = 1/BgPb[i]
    return EgTabel

def ZT(ZCor,Tres,TpcCorrection,P,PpcCorrection):
    ZTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        ZTabel[i] = Zmenu(ZCor,Tres,TpcCorrection,P[i],PpcCorrection)
    return ZTabel

def GasDensT(SgGas,T,P,Z):
    GasDensTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        GasDensTabel[i] = GasDensity(SgGas,T,P[i],Z[i])
    return GasDensTabel

def CwT(Cond,T,P,Rsw,TDS):
    CwTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        CwTabel[i] = CWater(Cond,T,P[i],Rsw[i],TDS)
    return CwTabel

def VisOilT(cond,mu0d,RsTabel,P,Pb,VisOilPb):
    OilViscosityTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(12):
        OilViscosityTabel[i] = mu0(cond,mu0d,RsTabel[i],P[i],Pb,VisOilPb)
    for i in range(12,30):
        OilViscosityTabel[i] = mu0(cond,mu0d,RsTabel[i],P[i],Pb,VisOilPb)
    return CwTabel

def CgT(ZCor,Tres,TPCCorrection,PTabel,PPCCorrection):
    CgTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        CgTabel[i] = Cg(ZCor,Tres,TPCCorrection,PTabel[i],PPCCorrection)
    return CgTabel

def GasVisT(cond,Tpc,Ppc,CO2,N2,H2S,SgGas,Tres,PTabel,GasDensityTabel):
    GasVisTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        GasVisTabel[i] = GasViscosity(cond,Tpc,Ppc,CO2,N2,H2S,SgGas,Tres,PTabel[i],GasDensityTabel[i])
    return GasVisTabel

def CoT(CondSat,CondUsat,RsTabel,P,Tres,API,SgGas,Pb,BgTabel,BoTabel,Psep,Tsep):
    CoTabel = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    for i in range(31):
        CoTabel[i] = Co(CondSat,CondUsat,RsTabel[i]*1000,RsTabel[12]*1000,P[i],Tres,API,SgGas,Pb,BgTabel[i]/1000,BoTabel[i],Psep,Tsep)
    return CoTabel
#########################################################################################################################################
PTabel = ()
SatCondTabel = ()
RsTabel = ()
BoTabel = ()
RsTabel = ()
BgTabel = ()
EgTabel = ()
BwTabel = ()
RswTabel = ()
H20GasTabel = ()
BrineDensityTabel = ()
BrineViscosityTabel = ()
ZTabel = ()
GasDensTabel = ()
CwTabel = ()


###########################################################################################

Hapus()
print('Apakah file berada pada direktori D:\PVT.xlsx?\n1. Ya\n2. Tidak')
pilihanFile = int(input('Masukkan pilihan anda: '))
Hapus()
if pilihanFile == 1:
    try:
        workbook = openpyxl.load_workbook("D:\PVT.xlsx")
        S = False
    except:
        print('Lokasi file tidak sesuai')
elif pilihanFile == 2:
    namaFile = input('Masukkan direktori file, beserta dengan ekstensinya\nContoh : D:\PVT.xlsx\nMasukkan direktori: ')
    workbook = openpyxl.load_workbook(namaFile)

sheet = workbook.active
Hapus()
while T:
    Hapus()
    PilihanUser = Tampilan()
    # Hasil PVT Calculator
    if PilihanUser==1:
        Preservoir = float(input('Masukkan tekanan Reservoir (psia): '))
        Hapus()
        DataGeneral = GeneralData()
        DataOil = OilData(DataGeneral[0],DataGeneral[3])
        DataImpurities = Impurities()
        DataBrine = BrineData()
        #KORELASI
        CorrelationTPCPPC = TPCPPCCorrelation()
        CorrectionCritProp = CorrectionCriticalProperties()
        ZCor = ZCorrelation()
        GasVisCor = GasViscosityCorrelation()
        DataKorelasiSaturated = KorelasiSaturated() #Tinjau lagi penggunaannya
        DataKorelasiUndersaturated = KorelasiUndersaturated() #Tinjau lagi penggunaannya
        DataBrineViscosityCorrelation = BrineViscosityCorrelation()
        HasilRsPb = Rs(DataKorelasiSaturated[1],Preservoir,DataGeneral[3],DataGeneral[0],DataOil[0],DataOil[1],DataOil[2],DataOil[3])
        HasilBoPb = Bo(DataKorelasiSaturated[0],DataKorelasiUndersaturated[0],HasilRsPb,DataGeneral[3],DataOil[0],DataGeneral[0],DataOil[3],DataOil[3],HasilRsPb,DataOil[1],DataOil[2],HasilBoPb)
    elif PilihanUser ==2:
        HasilRs = Rs(DataKorelasiSaturated[1],Preservoir,DataGeneral[3],DataGeneral[0],DataOil[0],DataOil[1],DataOil[2],DataOil[3])
        HasilBo = Bo(DataKorelasiSaturated[0], DataKorelasiUndersaturated[0],HasilRs,DataGeneral[3], DataOil[0], DataGeneral[0], Preservoir, DataOil[3], HasilRsPb, DataOil[1],DataOil[2] , HasilBoPb)#TINJAU LAGI TABEL 21
        Hasilrho0 = OilDens(DataKorelasiSaturated[2],DataKorelasiUndersaturated[1],DataOil[0],HasilRs*1000,HasilRsPb*10000,HasilBo,HasilBoPb,DataGeneral[3],Preservoir,DataOil[3],HasilBoPb,DataGeneral[0],DataOil[2],DataOil[1],36.2592)
        Hasilmu0d = mu0d(DataKorelasiSaturated[4], DataOil[0], DataGeneral[0])
        HasilVisOilPb = mu0(DataKorelasiSaturated[3],Hasilmu0d,HasilRsPb,Preservoir,DataOil[3],DataOil[3])
        Hasilmu0 = mu0(DataKorelasiSaturated[3],Hasilmu0d,HasilRs,Preservoir,DataOil[3],HasilVisOilPb)
        OilViscosityTabel[12] = Hasilmu0
        HasilVisOilPb = mu0(DataKorelasiSaturated[3],Hasilmu0d,HasilRsPb,Preservoir,DataOil[3],DataOil[3])
        HasilCo = Co(DataKorelasiSaturated[5],DataKorelasiUndersaturated[2],HasilRs*1000,HasilRsPb*1000,Preservoir,DataGeneral[0],DataOil[0],DataGeneral[3],DataOil[3],HasilBg/1000,HasilBo,DataOil[1],DataOil[2])
        print('OIL PROPERTIES')
        Conditions(Preservoir, DataOil[3])
        print(f'Bo : {HasilBo:.4f} RBL/STB')
        print(f'Rs : {HasilRs:.4f} Mscf/STB')
        print(f'ρ0 : {Hasilrho0:.4f} lbm/cf')
        print(f'μ0 : {Hasilmu0:.4f} cp')
        print(f'μ0d : {Hasilmu0d:.4f} cp')
        print(f'Co : {HasilCo:.4f} 1/MMpsi')
        Garis()

        print('CRITICAL PROPERTIES')
        HasilTpc = Tpc(CorrelationTPCPPC,DataGeneral[3])
        print(f'Tpc : {HasilTpc:.4f}°R')
        HasilPpc = Ppc(CorrelationTPCPPC,DataGeneral[3])
        print(f'Ppc : {HasilPpc:.4f} psia')
        HasilTpcCorrection = TpcCorrection(CorrectionCritProp,1,Tpc(CorrelationTPCPPC,DataGeneral[3]),Ppc(CorrelationTPCPPC,DataGeneral[3]),DataImpurities[1],DataImpurities[0],DataImpurities[2])
        print(f"Tpc : {HasilTpcCorrection:.2f}°R")
        HasilPpcCorrection = TpcCorrection(CorrectionCritProp,2,HasilTpc,HasilPpc,DataImpurities[1],DataImpurities[0],DataImpurities[2])
        print(f"Ppc : {HasilPpcCorrection:.2f} psia")
        HasilTpr = (DataGeneral[0])/TpcCorrection(CorrectionCritProp,1,Tpc(CorrelationTPCPPC,DataGeneral[3]),Ppc(CorrelationTPCPPC,DataGeneral[3]),DataImpurities[1],DataImpurities[0],DataImpurities[2])
        print(f'Tpr : {HasilTpr:.2f}')
        HasilPpr = Preservoir/TpcCorrection(CorrectionCritProp,2,Tpc(CorrelationTPCPPC,DataGeneral[3]),Ppc(CorrelationTPCPPC,DataGeneral[3]),DataImpurities[1],DataImpurities[0],DataImpurities[2])
        print(f"Ppr : {HasilPpr:.2f}")
        Garis()

        print('GAS PROPERTIES')
        HasilZFactor = ZFactor(ZCor, HasilTpr, HasilPpr)
        print(f'Z Factor : {HasilZFactor:.4f}')
        HasilBg = ((0.005035 * int(DataGeneral[0]) * 1000 * HasilZFactor/Preservoir  ))
        print(f'Bg : {HasilBg:.4f}  RB/Mscf')
        HasilH20 = H20inGas(DataBrine[1],DataGeneral[0],Preservoir)
        print(f'H2O in Gas : {HasilH20:.4f} lbm/Mscf')
        HasilGasDensity = GasDensity(DataGeneral[3],DataGeneral[0], Preservoir,HasilZFactor)
        HasilGasViscosity = GasViscosity(GasVisCor,HasilTpcCorrection, HasilPpcCorrection,DataImpurities[0],DataImpurities[2],DataImpurities[1],DataGeneral[3],DataGeneral[0],Preservoir,HasilGasDensity)
        print(f'Gas Viscosity : {HasilGasViscosity:.4f} cp')
        HasilGasDensity = GasDensity(DataGeneral[3],DataGeneral[0], Preservoir,HasilZFactor)
        print(f'Gas density : {HasilGasDensity:.4f} lbm/cf')
        HasilCg = Cg(ZCor,DataGeneral[0],HasilTpcCorrection,Preservoir,HasilPpcCorrection)
        print(f'Cg : {HasilCg:.4f}/MMpsi')
        print(f'Eg : {1/HasilBg:.4f} Mscf/Rb')
        Garis()
        print('BRINE PROPERTIES')
        HasilBw = Bw(DataBrine[0],DataGeneral[0],Preservoir)
        print(f'Bw : {HasilBw:.4f} rbl/STB')
        HasilRsw = Rsw(DataGeneral[0],Preservoir,DataBrine[1])
        print(f'Rsw : {HasilRsw:.4f} scf/STB')
        HasilBrineDensity = BrineDensity(DataBrine[1],HasilBw)
        print(f'Brine Density : {HasilBrineDensity:.4f} lbm/cf')
        HasilBrineViscosity = BrineViscosity(DataBrineViscosityCorrelation,DataBrine[1],Preservoir,DataGeneral[0])
        print(f'Brine Viscosity: {HasilBrineViscosity:.4f} cp')
        HasilCwater = CWater(DataBrine[0],DataGeneral[0],Preservoir,HasilRsw,DataBrine[1])
        print(f'Cwater : {HasilCwater:.4f} /MMpsi')
        Garis()
        ##################
        PTabel = P(DataGeneral[1],DataOil[3],DataGeneral[2])
        SatCondTabel = SatCond(PTabel,DataOil[3])
        RsTabel = Rsb(DataKorelasiSaturated[1],PTabel,DataGeneral[3],DataGeneral[0],DataOil[0],DataOil[1],DataOil[2],DataOil[3])
        RsTabel = Rsb(DataKorelasiSaturated[1],PTabel,DataGeneral[3],DataGeneral[0],DataOil[0],DataOil[1],DataOil[2],DataOil[3])
        BoTabel = BoPb(DataKorelasiSaturated[0],DataKorelasiUndersaturated[0],RsTabel,DataGeneral[3],DataOil[0],DataGeneral[0],PTabel,DataOil[3],RsTabel[12],DataOil[1],DataOil[2],1.1)
        BwTabel = BwPb(DataBrine[0],DataGeneral[0],PTabel)
        RswTabel = RswPb(DataGeneral[0],PTabel,DataBrine[1])
        H20GasTabel = H20Gas(DataBrine[1],DataGeneral[0],PTabel)
        BrineDensityTabel = BrineDensityT(DataBrine[1],BwTabel)
        BrineViscosityTabel = BrineViscosityT(DataBrineViscosityCorrelation,DataBrine[1],PTabel,DataGeneral[0])
        ZTabel = ZT(ZCor,DataGeneral[0],HasilTpcCorrection,PTabel,HasilPpcCorrection)
        GasDensTabel =  GasDensT(DataGeneral[3],DataGeneral[0],PTabel,ZTabel)
        BgTabel = BgT(ZTabel,DataGeneral[0],PTabel)
        EgTabel = EgT(BgTabel)
        CwTabel = CwT(DataBrine[0],DataGeneral[0],PTabel,RswTabel,DataBrine[1])
        OilViscosityTabel = VisOilT(DataKorelasiSaturated[3],Hasilmu0d,RsTabel,PTabel,DataOil[3],HasilVisOilPb)
        CgTabel = CgT(ZCor,DataGeneral[0],HasilTpcCorrection,PTabel,HasilPpcCorrection)
        GasVisTabel = GasVisT(GasVisCor,HasilTpcCorrection,HasilPpcCorrection,DataImpurities[0],DataImpurities[2],DataImpurities[1],DataGeneral[3],DataGeneral[0],PTabel,GasDensTabel)
        CoTabel = CoT(DataKorelasiSaturated[5],DataKorelasiUndersaturated[2],RsTabel,PTabel,DataGeneral[0],DataOil[0],DataGeneral[3],DataOil[3],BgTabel,BoTabel,DataOil[1],DataOil[2])
        #HasilCoPb = Co(DataKorelasiSaturated[5],DataKorelasiUndersaturated[2],HasilRsPb*1000,HasilRsPb*1000,)
        #HasilRho0Pb = OilDens(DataKorelasiSaturated[2],DataKorelasiUndersaturated[1],DataOil[0],HasilRsPb*1000,HasilRsPb*1000,HasilBoPb,HasilCoPb,DataGeneral[3],)
        inputAkhirProperties = int(input('Apakah ingin melanjutkan program?\n1. Ya\n2. Tidak\nMasukkan pilihan anda: '))
        Hapus()
        if inputAkhirProperties == 1:
            T = True
        elif inputAkhirProperties == 2:
            inputFinalProperties = int(input('Apakah anda yakin ingin keluar dari program?\n 1. Ya\n 2. Tidak\n Masukkan pilihan anda: '))
            if inputFinalProperties == 1:
                T = False
                Hapus()
                print('Terimakasih telah menggunakan program ini')
            elif inputFinalProperties == 2:
                T = True
                Hapus()
    elif PilihanUser == 3:
        if len(str(DataGeneral)) == 0 or len(str(DataOil)) == 0 or len(str(DataImpurities)) == 0 or len(str(DataBrine)) == 0 or len(str(CorrelationTPCPPC)) == 0:
            print('Data masih kosong, apakah ingin lanjut isi data dahulu \n 1. Ya\n 2. Tidak')
            PilihanUser1 = int(input('Masukkan Pilihan: '))
            if PilihanUser1 == 1:
                T = True
            elif PilihanUser1 == 2:
                print('Terimakasih telah menggunakan software ini')
                T = False
        else:    
            print('Data apa yang ingin dirubah?\n1. General Data\n2. Oil Data\n3. Impurities Data\n4. Brine Data\n5. Korelasi TPC dan PPC\n6. Korelasi oil pada kondisi Saturated\n7. Korelasi oil pada kondisi undersaturated\n8. Korelasi Gas Viscosity\n9. Korelasi Z-Factor\n10. Tekanan Reservoir\nNOTE : JANGAN LUPA SETELAH UBAH DATA, MASUK KE PVT CALCULATOR TERLEBIH DAHULU')
            Garis()
            PilihanPerubahan = int(input('Masukkan pilihan: '))
            Hapus()
            if PilihanPerubahan == 1:
                DataGeneral = GeneralData()
                Hapus()
            elif PilihanPerubahan == 2:
                DataOil = OilData(DataGeneral[0],DataGeneral[3])
                Hapus()
            elif PilihanPerubahan == 3:
                DataImpurities = Impurities()
                Hapus()
            elif PilihanPerubahan == 4:
                DataBrine = BrineData()
                Hapus()
            elif PilihanPerubahan == 5:
                CorrelationTPCPPC = TPCPPCCorrelation()
                Hapus()
            elif PilihanPerubahan == 6:
                DataKorelasiSaturated = KorelasiSaturated()
                Hapus()
            elif PilihanPerubahan == 7:
                DataKorelasiUndersaturated = KorelasiUndersaturated()
                Hapus()
            elif PilihanPerubahan == 8:
                GasVisCor = GasViscosityCorrelation()
                Hapus()
            elif PilihanPerubahan == 9:
                ZCor = ZCorrelation()
                Hapus()
            elif PilihanPerubahan == 10:
                Preservoir = float(input('Masukkan tekanan Reservoir (psia): '))
                Hapus()
            
    elif PilihanUser == 4:
        Hapus()
        print(f'Tekanan Reservoir (psia): {Preservoir}')
        Garis()
        LihatData(DataGeneral, DataOil, DataImpurities,DataBrine,CorrelationTPCPPC,CorrectionCritProp,ZCor,GasVisCor)
        Garis()
        inputLihatData = int(input('Apakah ingin lanjut program?\n 1. Ya\n 2. Tidak\n Masukkan Pilihan: '))
        Hapus()
    elif PilihanUser == 5:
        T = False
        Garis()
        print('TERIMAKASIH TELAH MENGGUNAKAN SOFTWARE INI')
        Garis()
    elif PilihanUser == 6:
        Hapus()
        print('CETAK GRAFIK\n1. P vs Bw\n2. P vs Bo\n3. P vs Rs\n4. P vs Rsw\n5. P vs Water Content\n6. P vs Brine Density\n7. P vs Brine Viscosity\n8. P vs Z\n9. P vs Bg\n10. P vs Eg\n11. P vs Cg\n12. P vs Gas Viscosity\n13. P vs Co')
        pilihanUserGrafik = int(input('Masukkan pilihan: '))
        Hapus()
        if pilihanUserGrafik == 1:
            T = TampilGrafik(PTabel,BwTabel,DataOil[3])
        elif pilihanUserGrafik == 2:
            T = TampilGrafik(PTabel,BoTabel,DataOil[3])
        elif pilihanUserGrafik == 3:
            T = TampilGrafik(PTabel,RsTabel,DataOil[3])
        elif pilihanUserGrafik == 4:
            T = TampilGrafik(PTabel,RswTabel,DataOil[3])
        elif pilihanUserGrafik == 5:
            T = TampilGrafik(PTabel,H20GasTabel,DataOil[3])
        elif pilihanUserGrafik == 6:
            T = TampilGrafik(PTabel,BrineDensityTabel,DataOil[3])
        elif pilihanUserGrafik == 7:
            T = TampilGrafik(PTabel,BrineViscosityTabel,DataOil[3])
        elif pilihanUserGrafik == 8:
            T = TampilGrafik(PTabel,ZTabel,DataOil[3])
        elif pilihanUserGrafik == 9:
            T = TampilGrafik(PTabel,BgTabel,DataOil[3])
        elif pilihanUserGrafik == 10:
            T = TampilGrafik(PTabel,EgTabel,DataOil[3]) 
        elif pilihanUserGrafik == 11:
            T = TampilGrafik(PTabel,CgTabel,DataOil[3])
        elif pilihanUserGrafik == 12:
            T = TampilGrafik(PTabel,GasVisTabel,DataOil[3])
        elif pilihanUserGrafik == 13:
            T = TampilGrafik(PTabel,CoTabel,DataOil[3])
    elif PilihanUser == 7:
        # INPUT P TABEL
        for i in range(31):
            sheet[f'A{i+3}'] = PTabel[i]
        # INPUT CONSAT TABEL
        for i in range(31):
            sheet[f'B{i+3}'] = SatCondTabel[i]
        # INPUT Rs TABEL
        for i in range(31):
            sheet[f'D{i+3}'] = RsTabel[i]
        # INPUT Bo TABEL
        for i in range(31):
            sheet[f'C{i+3}'] = BoTabel[i]
        # INPUT Bg TABEL
        for i in range(31):
            sheet[f'E{i+3}'] = BgTabel[i]
        # INPUT Eg TABEL
        for i in range(31):
            sheet[f'F{i+3}'] = EgTabel[i]
        # INPUT Bw TABEL
        for i in range(31):
            sheet[f'G{i+3}'] = BwTabel[i]
        # INPUT Rsw TABEL
        for i in range(31):
            sheet[f'H{i+3}'] = RswTabel[i]
        # INPUT H20 Content TABEL
        for i in range(31):
            sheet[f'I{i+3}'] = H20GasTabel[i]
        # INPUT ZFactor TABEL
        for i in range(31):
            sheet[f'J{i+3}'] = ZTabel[i]
        # INPUT GasDensity TABEL
        for i in range(31):
            sheet[f'L{i+3}'] = GasDensTabel[i]
        # INPUT BrineDens TABEL
        for i in range(31):
            sheet[f'M{i+3}'] = BrineDensityTabel[i]
        # INPUT OilVis TABEL
        for i in range(31):
            sheet[f'N{i+3}'] = OilViscosityTabel[i]
        # INPUT Gas Viscosity TABEL
        for i in range(31):
            sheet[f'O{i+3}'] = GasVisTabel[i]
        # INPUT Brine Viscosity TABEL
        for i in range(31):
            sheet[f'P{i+3}'] = BrineViscosityTabel[i]
        # INPUT Co TABEL
        for i in range(31):
            sheet[f'Q{i+3}'] = CoTabel[i]
        # INPUT Cg TABEL
        for i in range(31):
            sheet[f'R{i+3}'] = CgTabel[i]
        # INPUT Cw TABEL
        for i in range(31):
            sheet[f'S{i+3}'] = CwTabel[i]
        if pilihanFile == 1:
            workbook.save('D:\\PVT.xlsx')
        elif pilihanFile == 2:
            workbook.save(namaFile)
        T = True
    elif type(PilihanUser) == "<class 'str'>":
        Hapus()
        inputSalah = input('anda salah input, apakah ingin lanjut program?\n1. Ya\n2. Tidak\nMasukkan pilihan anda: ')
        if inputSalah == 1 or inputSalah == '1':
            T = True
        elif inputSalah == 2 or inputSalah == '2':
            T = False
    else:
        Hapus()
        inputSalah = input('anda salah input, apakah ingin lanjut program?\n1. Ya\n2. Tidak\nMasukkan pilihan anda: ')
        if inputSalah == 1 or inputSalah == '1':
            T = True
        elif inputSalah == 2 or inputSalah == '2':
            T = False

        


