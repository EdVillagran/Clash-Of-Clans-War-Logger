import requests
import xlrd
import openpyxl
import clash_headers as ch

#Multiple functions need access to API info
#access_API returns dict with the JSON info stored
def access_API():
    url_in_War="https://api.clashofclans.com/v1/clans/%23GV28G8Q/currentwar"
    headers=ch.heading()    #Contains authorization info
    response=requests.get(url_in_War, headers=headers)
    resp_json=response.json()

    return resp_json

def update_Sheet(players):
    #Access Excel sheet
    loc2=("Documents/WEE_ONES.xlsx")
    wb = openpyxl.load_workbook(loc2)
    sheet=wb.get_sheet_by_name('Sheet1')

    #Variables to define current cell
    letter='A'
    number=3
    cell=letter+str(number)

    #Find all names currently in Excel sheet:
    namelist=[]
    while sheet[cell].value != None:
        namelist.append(sheet[cell].value)
        number+=1
        cell=letter+str(number)


    for player in players:
        # Excel Columns: A             B       C  D  E  F  G  H  I
        #List Value:   ['Mmoc', '#RJ9PCVR8', 79, 2, 0, 2, 1, 1, 0]
        #Index:           0       1          2   3  4  5  6  7  8

        #Create variables for readability
        name=player[0]
        tag=player[1]
        avgDestruction=player[2]
        stars=player[3]
        Successhit=player[4]
        avgstars=player[5]
        missed=player[6]

        if player[0] in namelist:
            #Create the index number from excel row
            number=namelist.index(player[0])+3

            print("Updating Player:",player[0])
            #Update Values in excel sheet

            sheet['C'+str(number)]=(sheet['C'+str(number)].value+avgDestruction)/2
            sheet['D'+str(number)]=sheet['D'+str(number)].value+stars
            sheet['E'+str(number)]=sheet['E'+str(number)].value+Successhit
            sheet['F'+str(number)]=(sheet['F'+str(number)].value+avgstars)/2
            sheet['G'+str(number)]=sheet['G'+str(number)].value+missed
            sheet['H'+str(number)]=sheet['H'+str(number)].value+1
        else:
            print(name, "Not in sheet, Adding")
            newrow=len(namelist)+3      #New row to add them too
            namelist.append(player[0])  #Add names to those who are on sheet

            #Update Values in excel sheet
            sheet['A'+str(newrow)]=name
            sheet['B'+str(newrow)]=tag
            sheet['C'+str(newrow)]=avgDestruction
            sheet['D'+str(newrow)]=stars
            sheet['E'+str(newrow)]=Successhit
            sheet['F'+str(newrow)]=avgstars
            sheet['G'+str(newrow)]=missed
            sheet['H'+str(newrow)]=1    #War counter
            sheet['I'+str(newrow)]=0    #Not in use yet

    #Save Workbook
    wb.save(loc2)

def WarStats():

    resp_json=access_API()
    store=[]
    for players in resp_json['clan']['members']:

        for data in players:

            #If in war show attack field
            if data=='attacks':

                for attack in players[data]:

                    tag=attack['attackerTag']

                    stars=attack['stars']

                    if stars==3:
                        success=1
                    else:
                        success=0

                    destruction=int(attack['destructionPercentage'])
                    store.append([tag,destruction,stars,success])


    for player in store:
        index=store.index(player)

        for player2 in store[index+1:]:
            if (player[0]==player2[0]):

                index2=store.index(player2)

                #[tag,destruction,stars,success]
                #Avg percentage
                player[1]=(player[1]+player2[1])/2

                player.append(int(str((player[2]+player2[2])/2)[0]))

                #Total Stars
                player[2]+=player2[2]

                #Total Perfect Hits
                player[3]+=player2[3]
                #Avg star per hit

                player.append(0)
                store.pop(index2)

    #Count for any missed attacks
    for i in store:

        if len(i)==4:
            i.append(i[2])
            i.append(1)

    return store
#stores all data on each player into each players list
def build_player(names,stats):
    #List to store each into
    built=[]
    for attack in stats:
        for player in names:
            if attack[0]==player[1]:

                #built = [name, tag, avgDes,total stars, perfect hits, total missed, war count=1]
                built.append([player[0],player[1],attack[1],attack[2],attack[3],attack[4],attack[5],1])
                
    return built

def main():


    resp_json = access_API()
    stats=WarStats()

    #Gather name on every player in current war
    member_stats = resp_json['clan']['members']
    players=[]

    for row in member_stats:
        name=row['name']
        tag=row['tag']
        players.append([name,tag])

    finished_list=build_player(players,stats)

    #For Chance to 3 star
    for i in finished_list:
        i.append(0)

    update_Sheet(finished_list)

main()
