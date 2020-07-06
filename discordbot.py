import discord
from discord.ext import commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import os

#スプレッドシート情報欄
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('test-python-275700-5f70ae0934df.json', scope)
gc = gspread.authorize(credentials)
SPREADSHEET_KEY = "1t1jtRXdBAWPi-xm21LhZ61CioOyzFaNChBNbud2V1lg"

wkb = gc.open_by_key(SPREADSHEET_KEY)
wks1 = wkb.worksheet("ID検索") #ID検索
wks2 = wkb.worksheet("操作用シート")
wks3 = wkb.worksheet("退団ID")
wks4 = wkb.worksheet("退団理由")

#discord情報欄
bot = commands.Bot(command_prefix='!')
token = os.environ['TOKEN']

dt_now = datetime.datetime.now()

def get_List():
    data = []
    rank = wks1.range("B4:B53")
    name = wks1.range("C4:C53")
    ID = wks1.range("D4:D53")
    MS = wks1.range("E4:E53")
    Active = wks1.range("F4:F53")

    for i in range(len(rank)):
        if Active[i].value == "活動中":
            Active[i] = ""
        else:
            Active[i] = "休止中(" + Active[i].value + ")"
        a = {"Rank":rank[i].value,"Name":name[i].value,"ID":ID[i].value,"MaxStage":MS[i].value,"Active":Active[i]}
        data.append(a)
    return data

@bot.event
async def on_ready():
    print('Logged in as')
    print(bot.user.name)
    print(bot.user.id)
    print('------')

# -------------------------------------------------------------------------------------------------------------
# コマンド関連
# -------------------------------------------------------------------------------------------------------------

### menber list表示
@bot.command()
async def ml(ctx):    
    data = get_List()
    
    m = "**MEMBER LIST**\n```\n"
    m += "{0:^4}|{1:^8}".format("Rank","Name") + "\n"
    for i in data:
        m += "{Rank:^4}|{Name:<} {Active}".format(**i) + "\n"
    m += "```"
    await ctx.send(m)
# -------------------------------------------------------------------------------------------------------------

### info表示
@bot.command()
async def info(ctx,*args):
    data = get_List()
    
    embed=discord.Embed(title="__***Player Info***__",color=0xee1111)
    embed.add_field(name="Name", value=data[int(args[0])-1]["Name"], inline=False)
    embed.add_field(name="ID", value=data[int(args[0])-1]["ID"], inline=True)
    embed.add_field(name="MaxStage", value=data[int(args[0])-1]["MaxStage"], inline=True)
    embed.add_field(name="Rejoin", value=str(data[int(args[0])-1]["Name"].count("＊")) + "回目", inline=True)
    if data[int(args[0])-1]["Name"].count("＊") == 0:
        embed.add_field(name="Reason for leaving", value="なし", inline=True)
    else:
        search_ID = wks4.range("A1:A1000")
        for i in range(len(search_ID)):
            search_ID[i] = search_ID[i].value
            if search_ID[i] == data[int(args[0])-1]["ID"]:
                num = i + 1
                break
        
        m = ""
        for i in range(25):
            value = wks4.cell(num,2+i).value
            if value == "":
                break
            m = m + value + "\n"

        embed.add_field(name="Reason for leaving", value=m, inline=False)

    await ctx.send(embed=embed)
# -------------------------------------------------------------------------------------------------------------

### MS update
@bot.command()
async def MSupdate(ctx,*args):
    data = get_List()
    
    # A:name , B:ID , C:MS
    up_ID = wks2.range("B1:B1000")
    for i in range(len(up_ID)):
        up_ID[i] = up_ID[i].value
        if up_ID[i] == data[int(args[0])-1]["ID"]:
            break
    num = up_ID.index(data[int(args[0])-1]["ID"])
    wks2.update_cell(num+1,3,int(args[1]))

    m = "```\n" +\
        data[int(args[0])-1]["Name"] + "のMSを更新しました\n" +\
        "```" 
    await ctx.send(m)
# -------------------------------------------------------------------------------------------------------------

### 休止状態に変更
@bot.command()
async def sleep(ctx,*args):
    data = get_List()

    # A:name , B:ID , C:MS
    up_ID = wks2.range("B1:B1000")
    for i in range(len(up_ID)):
        up_ID[i] = up_ID[i].value
        if up_ID[i] == data[int(args[0])-1]["ID"]:
            break
    num = up_ID.index(data[int(args[0])-1]["ID"])
    wks2.update_cell(num+1,4,dt_now.strftime('%Y/%m/%d %H:%M:%S'))
    
    day = dt_now + datetime.timedelta(days=21)
    day.strftime('%Y/%m/%d')

    m = "```\n" +\
        data[int(args[0])-1]["Name"] + "を休止状態にしました\n" +\
        "終了期限は" + day.strftime('%Y/%m/%d') + "です\n" +\
        "```" 
    await ctx.send(m)
# -------------------------------------------------------------------------------------------------------------

### 活動状態に変更
@bot.command()
async def active(ctx,*args):
    data = get_List()

    # A:name , B:ID , C:MS
    up_ID = wks2.range("B1:B1000")
    for i in range(len(up_ID)):
        up_ID[i] = up_ID[i].value
        if up_ID[i] == data[int(args[0])-1]["ID"]:
            break
    num = up_ID.index(data[int(args[0])-1]["ID"])
    wks2.update_cell(num+1,4,"活動中")

    m = "```\n" +\
        data[int(args[0])-1]["Name"] + "を活動状態にしました\n" +\
        "```" 
    await ctx.send(m)
# -------------------------------------------------------------------------------------------------------------

### メンバー削除
@bot.command()
async def delete(ctx,*args):
    if len(args) == 2:
        data = get_List()
        
        # A:name , B:ID , C:MS
        up_ID = wks2.range("B1:B1000")
        for i in range(len(up_ID)):
            up_ID[i] = up_ID[i].value
            if up_ID[i] == data[int(args[0])-1]["ID"]:
                break
        num = up_ID.index(data[int(args[0])-1]["ID"])
        for i in range(4):
            wks2.update_cell(num+1,2+i,"")
        
        out_ID = wks3.range("A1:A1000")
        for i in range(len(out_ID)):
            if out_ID[i].value == "":
                pos = i + 1
                break
        wks3.update_cell(pos,1,data[int(args[0])-1]["ID"])
        wks3.update_cell(pos,2,args[1])

        search_ID = wks4.range("A1:A1000")
        for i in range(len(search_ID)):
            search_ID[i] = search_ID[i].value
            if search_ID[i] == data[int(args[0])-1]["ID"]:
                num = i + 1
                break
            elif search_ID[i] == "":
                break
        
        for i in range(25):
            value = wks4.cell(num,2+i).value
            if value == "":
                wks4.update_cell(num,2+i,args[1])
                break

        m = "```\n" +\
            data[int(args[0])-1]["Name"] + "を除名しました\n" +\
            "```" 
    else:
        m = "ランク、又は退団理由を記入してください"
    await ctx.send(m)
# -------------------------------------------------------------------------------------------------------------

### メンバー登録
@bot.command()
async def add(ctx,*args):
    if len(args) == 3:
        # A:name , B:ID , C:MS
        up_ID = wks2.range("B1:B1000")
        for i in range(len(up_ID)):
            if up_ID[i].value == "":
                pos = i + 1
                break
        wks2.update_cell(pos,2,args[1])
        wks2.update_cell(pos,3,args[2])
        wks2.update_cell(pos,4,"活動中")
        wks2.update_cell(pos,5,args[0])
    m = "```\n" +\
        args[0] + "を登録しました\n" +\
        "```" 
    await ctx.send(m)
# -------------------------------------------------------------------------------------------------------------

### help表示
@bot.command()
async def hlp(ctx,*args):
    m = "```\n" +\
        "**コマンド一覧**\n" +\
        "!ml\n" +\
        "> メンバーリストの表示\n" +\
        "\n" +\
        "!info <rank>\n" +\
        "> 名前、ID、MSの表示\n" +\
        "\n" +\
        "!MSupdate <rank> <MaxStage>\n" +\
        "> MSの更新\n" +\
        "\n" +\
        "!sleep <rank>\n" +\
        "> 休止状態に変更\n" +\
        "\n" +\
        "!active <rank>\n" +\
        "> 活動状態に変更\n" +\
        "\n" +\
        "!add <name> <ID> <MaxStage>\n" +\
        "> メンバーの登録\n" +\
        "\n" + \
        "!delete <rank> <reason>\n" +\
        "> メンバーの除名\n" +\
        "```"
    await ctx.send(m)
# -------------------------------------------------------------------------------------------------------------


bot.run(token)
