# *-* coding : utf- 8 *-*
import barcode
from barcode.writer import ImageWriter
from openpyxl import load_workbook

ac=load_workbook("aa.xlsx")
ac1=ac["Sayfa1"]

say=1
trend=0
hb=0
gg=0
n11=0

while say !=10:

    kod=ac1["A{}".format(say)].value
    alici=ac1["b{}".format(say)].value
    kod1=str(kod)
    kod2=[]

    for i in kod1:
        kod2.append(i)

    uzun=int(len(kod2))
    gondr=kod2 [0]

    if uzun ==13:
        if gondr =="6":
            ean = barcode.get('ean13', '{} '.format(kod), writer=ImageWriter())
            filename = ean.save('{}'.format(alici),text='HepsiBurada \n {} \n {}'.format(alici,kod))
            hb=hb+1
            say=say+1

        elif gondr == "7":
            ean = barcode.get('ean13', '{} '.format(kod), writer=ImageWriter())
            filename = ean.save('{}'.format(alici), text='Trendyol \n {} \n {}'.format(alici, kod))
            trend=trend+1
            say = say + 1

    elif uzun == 15 :
        ean = barcode.get('code128', '{} '.format(kod), writer=ImageWriter())
        filename = ean.save('{}'.format(alici), text='N11  \n {} \n {}'.format(alici, kod))
        n11 = n11 + 1
        say = say + 1

    elif uzun == 10:
        ean = barcode.get('code128', '{} '.format(kod), writer=ImageWriter())
        filename = ean.save('{}'.format(alici), text='GittiGidiyor \n {} \n {}'.format(alici, kod))
        gg = gg + 1
        say = say + 1

    else:
        print(hb, " Tane Hepsiburada oluşturuldu")
        print(trend, " Tane Trendyol oluşturuldu")
        print(gg, " Tane GittiGidiyor oluşturuldu")
        print(trend, " Tane N11 oluşturuldu")
        break

ac.close()