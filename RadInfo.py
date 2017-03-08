import datetime

class RadInfo(object):
    def __init__(this, name):
        this.name = name
        this.fileDates = []
        this.closedDates = []
        this.caseNames = []
        this.methodCloseds = []
        this.investigations = []
        this.disciplinaryActions = []
        this.longest = 0
        
    # Helpers
    def clearBlanks(this, entry):
        bad_entry = [None, u"", u"N/A", u"NA", u"N/A ", u"No", u"No ", u"0", u"N.A"]
        index = len(entry) - 1
        for i in range(index, -1, -1):
            if entry[i] in bad_entry:
                del entry[i]

    def convertDates(this, entry):
        index = len(entry)
        for i in range(index):
            entry[i] = datetime.datetime.strptime(entry[i], '%m/%d/%Y')
            
    # Setters     
    def setCaseNames(this, entry):
        #this.clearBlanks(entry)
        this.caseNames = entry
        if len(this.caseNames) > this.longest:
            this.longest = len(this.caseNames)

    def setFileDates(this, entry):
        this.clearBlanks(entry)
        this.convertDates(entry)
        this.fileDates = entry
        if len(this.fileDates) > this.longest:
            this.longest = len(this.fileDates)

    def setClosedDates(this, entry):
        this.clearBlanks(entry)
        this.convertDates(entry)
        this.closedDates = entry
        if len(this.closedDates) > this.longest:
            this.longest = len(this.closedDates)
            
    def setMethodCloseds(this, entry):
        #this.clearBlanks(entry)        
        this.methodCloseds = entry
        if len(this.methodCloseds) > this.longest:
            this.longest = len(this.methodCloseds)
            
    def setInvestigations(this, entry):
        this.clearBlanks(entry)        
        this.investigations = entry
        if len(this.investigations) > this.longest:
            this.longest = len(this.investigations)
            
    def setDisciplinaryActions(this, entry):
        this.clearBlanks(entry)
        this.disciplinaryActions = entry
        if len(this.disciplinaryActions) > this.longest:
            this.longest = len(this.disciplinaryActions)

    
    # Getters

    



        
