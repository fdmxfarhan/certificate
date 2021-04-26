import pygame
import xlrd

wb = xlrd.open_workbook('certificate.xls') 
sheet = wb.sheet_by_index(0)
num = int(input('Enter Number of columns: '))
scale = 0.3
# print (sheet.cell_value(0, 0))


pygame.init()
certificate_img = pygame.image.load('certificate.jpg')
rect = certificate_img.get_rect()
scaled_certificate_img = pygame.transform.scale(certificate_img, (int(rect[2] * scale), int(rect[3] * scale)))
rect2 = scaled_certificate_img.get_rect()
screen2 = pygame.display.set_mode((rect2[2], rect2[3]))

black = (101,81,68)
white = (255,255,255)
red = (255,0,0)
fontSize1 = 80
fontSize2 = 150
fontSize3 = 120

class Info:
    def __init__(self):
        self.idNumber = ''
        self.englishFirstName = ''
        self.englishLastName = ''
        self.englishLeague = ''
        self.englishAffiliation = ''
        self.teamName = ''
        self.nameX = 1772
        self.nameY = 1392
        self.englishLeagueX = 1783
        self.englishLeagueY = 1686
        self.englishAffiliationX = 2613
        self.englishAffiliationY = 1686
        self.teamNameX = 1076
        self.teamNameY = 1686
    def printInfo(self):
        print(
            'idNumber: ' + str(self.idNumber) + '\n' + 
            'englishFirstName: ' + str(self.englishFirstName) + '\n' + 
            'englishLastName: ' + str(self.englishLastName) + '\n' + 
            'englishLeague: ' + str(self.englishLeague) + '\n' + 
            'englishAffiliation: ' + str(self.englishAffiliation) + '\n' + 
            'teamName: ' + str(self.teamName) + '\n' + 
            'nameX: ' + str(self.nameX) + '\n' + 
            'nameY: ' + str(self.nameY) + '\n' + 
            'englishLeagueX: ' + str(self.englishLeagueX) + '\n' + 
            'englishLeagueY: ' + str(self.englishLeagueY) + '\n' + 
            'englishAffiliationX: ' + str(self.englishAffiliationX) + '\n' + 
            'englishAffiliationY: ' + str(self.englishAffiliationY) + '\n' + 
            'teamNameX: ' + str(self.teamNameX) + '\n' + 
            'teamNameY: ' + str(self.teamNameY) + '\n'
        )

info = [0] * num
for i in range(num):
    info[i] = Info()
    info[i].idNumber = sheet.cell_value(i, 0)
    info[i].englishFirstName = sheet.cell_value(i, 1)
    info[i].englishLastName = sheet.cell_value(i, 2)
    info[i].englishLeague = sheet.cell_value(i, 3)
    info[i].englishAffiliation = sheet.cell_value(i, 4)
    info[i].teamName = sheet.cell_value(i, 5)
    # info[i].printInfo()

ITCEDSCR = pygame.font.Font('ITCEDSCR.TTF',fontSize1)
SacramentoRegular = pygame.font.Font('Sacramento-Regular.ttf',fontSize2)
font = pygame.font.Font('name.ttf',fontSize3)
ITCEDSCR2 = pygame.font.Font('ITCEDSCR.TTF',int(fontSize1 * scale))
SacramentoRegular2 = pygame.font.Font('Sacramento-Regular.ttf',int(fontSize2 * scale))
font2 = pygame.font.Font('name.ttf',int(fontSize3 * scale))

# done = False
# while not done:
#     for event in pygame.event.get():
#         if event.type == pygame.QUIT:
#             done = True
#         if event.type == pygame.MOUSEBUTTONUP:
#             nameX = mouse[0] / scale
#             nameY = mouse[1] / scale
#             print(nameX, nameY)
#     mouse = pygame.mouse.get_pos()
    
#     screen2.blit(scaled_certificate_img, rect2)
#     nameText = font2.render('Name', True, black)
#     nameTextRect = nameText.get_rect()
#     screen2.blit(nameText, (mouse[0] - nameTextRect[2]/2, mouse[1] - nameTextRect[3]/2))

#     pygame.display.update()
    
screen = pygame.display.set_mode((rect[2], rect[3]))

for i in range(num):
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            crashed = True
    
    screen.blit(certificate_img, rect)
    nameText = SacramentoRegular.render(info[i].englishFirstName + ' ' + info[i].englishLastName, True, black)
    nameTextRect = nameText.get_rect()
    screen.blit(nameText, (info[i].nameX - nameTextRect[2]/2, info[i].nameY - nameTextRect[3]/2))
    
    nameText = ITCEDSCR.render(info[i].englishLeague, True, black)
    nameTextRect = nameText.get_rect()
    screen.blit(nameText, (info[i].englishLeagueX - nameTextRect[2]/2, info[i].englishLeagueY - nameTextRect[3]/2))
    
    nameText = ITCEDSCR.render(info[i].englishAffiliation, True, black)
    nameTextRect = nameText.get_rect()
    screen.blit(nameText, (info[i].englishAffiliationX - nameTextRect[2]/2, info[i].englishAffiliationY - nameTextRect[3]/2))
    
    nameText = ITCEDSCR.render(info[i].teamName, True, black)
    nameTextRect = nameText.get_rect()
    screen.blit(nameText, (info[i].teamNameX - nameTextRect[2]/2, info[i].teamNameY - nameTextRect[3]/2))
    
    
    pygame.image.save(screen, './certificates/' + sheet.cell_value(i, 0) + '.jpg')
    pygame.display.update()