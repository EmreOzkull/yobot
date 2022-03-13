import discord
import dispy_varlar
import random
from discord.ext import commands
from discord_ui import Button
from datetime import datetime
import time
from discord_components import DiscordComponents, ComponentsBot, Button, SelectOption, Select
import xlsxwriter
from openpyxl import load_workbook
import pandas as pd




ExcelDataInPandasDataFrame = pd.read_excel("projeson.xlsx")



TOKEN = dispy_varlar.TOKEN

client = discord.Client()
bot = commands.Bot(command_prefix='!')

with open('kotu_kelime.txt') as kufurler:
    kufurler = kufurler.read().split()


# Baglanilan Server'e ait Üye Bilgilerini Terminale aktar
@bot.event
async def on_ready():
    print(f'{bot.user.name} baglandi!')
    await bot.change_presence(activity=discord.Activity(type=discord.ActivityType.watching, name="Vox'un Kölesi"))
    for server in bot.get_all_members():
        print('Üye ve ID', end=":")
        print(f'{server}')

#Yeni Katilan Kullaniciya Hosgeldin Mesaji
@bot.event
async def on_member_join(member):
    await member.create_dm()
    await member.dm_channel.send(
        f'Yazılım Ofisi Discord Kanalina Hosgeldin {member.name}! '
    )


#Bilgilendirme Komutu
@bot.command(name='bilgi', help='Bot bilgilendirmesi')
async def bilgi(ctx):
    bilgi_cvp = "Bu Bot Yazılım Ofisi Discord Kanalinin Karsilama Botudur!"
    await ctx.send(bilgi_cvp)



# !zar-at denilince rastgele 2 zar ciktisi verir.
@bot.command(name='zar-at', help='Zar oyunu')
async def zar_at(ctx):
    a = random.choice(range(1,7))
    b = random.choice(range(1,7))
    zar = f'Sansina cikan zarlar: {a} ve {b}'
    await ctx.send(zar)

# !topla sayi1 sayi2 ile toplama islemi yapmasini saglar
@bot.command(name='topla',a=int,b=int,help="Toplama islemi yapmasini saglar")
async def topla(ctx,a:int,b:int):
    await ctx.send(a+b)

@bot.command()
async def sec(ctx, *sec:str):
    await ctx.send(random.choice(sec))

# !proje komutu ile yeni devam eden proje kanalı açar
@bot.command(name='proje', help='Devam eden proje kanalı açar')
@commands.has_role('🚀┋Crew')
async def proje(ctx, metin):
    proje_olusturma = bot.get_channel(947066750372028477)
    server = ctx.guild
    category = bot.get_channel(947065585483780166)
    if ctx.channel.id == (947066750372028477):
        await server.create_text_channel(f"「📂」{metin}",category=category)
        await ctx.send(f'**{ctx.author.mention} Proje Devam Kanalını Oluşturdum!\n Bu kanal üzerinden projenizin gelişim aşamaları, malzeme listesi ve proje raporlamasını yapmayı unutma!**',delete_after=20)

    else:
        await ctx.send(f"**{ctx.author.mention} Bu komutu yalnızca {proje_olusturma} kanalında kullanabiilirsin!**",delete_after=15)

# !proje komutu ile yeni devam eden proje kanalı açar
@bot.command(name='complete', help='Devam eden projeyi bitmiş projelere taşır.')
@commands.has_role('🔰┋Organizasyon Birimi')
async def complete(ctx):
    server = ctx.guild
    channel = ctx.channel
    category_ok = bot.get_channel(942322736091504701)
    await channel.edit(category=category_ok)
    await ctx.send(f'**{ctx.author.mention} Projeyi Tamamlanmış Projelere Taşıdım.\n !excell komutu ile projenin excell entegrasyonunu yapayı unutmayın!**',delete_after=20)

# !excell komutu ile projeyi excell e aktarır
@bot.command(name='excell', help='Projeyi excell e aktarır.')
@commands.has_role('🔰┋Organizasyon Birimi')
async def excell(ctx,member:discord.Member,proje_adi,link):
    server = ctx.guild
    member = str(member)
    proje_adi = proje_adi.title()
    proje_adi.replace("-"," ")

    if member == "niddahales#7520":
        ad = "?"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "adnaninci#9538":
        ad = "Adnan İnci"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "Basak 🦩#5459":
        ad = "Başak Yalçıner"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "betül#5070":
        ad = "Betül Altunel "
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "busem#8181":
        ad = "Busem Kanbaş"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "Egemen Bahtiyar#7646":
        ad = "Egemen Bahtiyar"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == 'Emre "voxventi"#0003':
        ad = "Emre Özkul"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "♆KAAN#6518":
        ad = "Kaan Kahraman"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "kemalKartal#3862":
        ad = "Kemal Kartal"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "kerimdw#4649":
        ad = "Kerim Dev"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "Levent_Ozdemir#5184":
        ad = "Levent Özdemir"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "Quasimodo💻#4471":
        ad = "Muhammed"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "salihyksl#8337":
        ad = "Salih Yüksel"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "Yarmax#0116":
        ad = "Süleyman Akıllı"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "TalhaNebiKumru#0897":
        ad = "Talha Nebi Kumru"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    elif member == "umutenss#6170":
        ad = "Umut Enes"
        ExcelDataInPandasDataFrame[member] = [ad, proje_adi,link, 1]
        ExcelDataInPandasDataFrame.to_excel("projeson.xlsx", index=False)

    await ctx.send(f"**{ctx.author.mention} Projeyi Excell'e Taşıdım.**",delete_after=20)


# !ses_kanali ile yeni ses kanali acar
@bot.command(name='ses_kanali', help='Ses Kanali olusturur')
@commands.has_role('🔰┋Organizasyon Birimi')
async def ses_kanali(ctx, ses):
    server = ctx.guild
    await server.create_voice_channel(ses)

# Administrator yetkisi olan kullanicilarin !kick kullanici_adi komutu ile kullanici atmasini saglar
@bot.command(name='kick', help= 'Birisini atmak icin kullan')
@commands.has_role('🔰┋Organizasyon Birimi')
async def kick(ctx, member: discord.Member, *, reason=None):
    await member.kick(reason=reason)
    await ctx.send(f'{member} Kullanıcısını sunucudan uzaklaştırdım.')

# Administrator yetkisi olan kullanicilarin !ban kullanici_adi komutu ile kullanici banlamasını sağlar
@bot.command(name='ban', help= 'Birisini banlamak icin kullan')
@commands.has_role('🔰┋Organizasyon Birimi')
async def ban(ctx, member: discord.Member, *, reason=None):
    await member.ban(reason=reason)
    await ctx.send(f'{member} Kullanıcısını sunucudan banladım.')


@bot.command(name='odaac', help= 'Oda açmak için kullan.')
async def odaac(ctx):
    await ctx.send("", components=[
        [Button(label="Oda Oluştur", style="3", emoji="✅", custom_id="button1"),
         Button(label="Oda Sil", style="4", emoji="❎", custom_id="button2")]
    ])
    interaction = await client.wait_for("button_click", check=lambda i: i.custom_id == "button1")
    await interaction.send(content="Button clicked!", ephemeral=False)

# Mesaj silme komutu
@bot.command(name='sil', help= 'Belirli sayıda mesajı silmek için kullanılır.')
async def sil(ctx, number:int):
    channel = ctx.channel
    messages = await channel.history(limit=number).flatten()
    try:
        await channel.delete_messages(messages)
        await ctx.send(f'**{number}** tane mesaj uzaya yollandı 🚀', delete_after=5)
    except:
        ctx.send("Bir ahtayla karşılaştın. Yavaşça Vox'un kulağına yaklaş ve fısılda!")



# Bot'un konusmalarda olusacak olan mesaj iceriklerine cevap vermesini saglar
@bot.event
async def on_message(message): # Botun hangi kanallardaki konusmalara bakmasi gerektigini belirler 

    if message.author == client.user:
        return

    selam_alma = ['Sa','Selam','s.a','sa','slm','selam']
    bot_selamlama = ['Aleyna Aleyküm esselam', 'Ooooooo kimleri görüyorum', 'Vaaaaaay Sen yasiyor muydun ya?','Vay a.s Haci cav cav , nörüyon?', 'A.s ne geziyon buralarda?','Selamina selam cigerim!','Bana selam verme selam tutmayi... Yok yok konular karisti.. A.s','Dedi naber dedim iyidir... A.s genc!','Merhabalar, Yazılım Ofisi Discord Kanalina Hosgeldiniz... NÖRÜYON Cigerparem?']



    if message.content.lower() in selam_alma:
        selam_ver = random.choice(bot_selamlama)
        await message.channel.send(selam_ver)



    elif 'özel ders' in message.content.lower():
        await message.channel.send('ssssss Ozel ders Konusu acilinca Saizzou sövüyor!')
        if message.author == client.user:
            return

    for kufur in kufurler:
        if kufur in message.content.lower():
            await message.delete()
            msg = f'{message.author.mention}! Kufur edeni tinkywinkylerim aklin basina gelir!'
            await message.channel.send(msg)


    await bot.process_commands(message)
bot.run(TOKEN)
