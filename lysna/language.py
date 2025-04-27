class English:
    """
    English language class for the LYSNA application.
    """

    def __init__(self):
        self.language = "English"
        self.language_code = "en"
        
        self.scales = {
            "org_core": "Organisational Core",
            "team_core": "Team Core",
            "safety": "Psychological Safety",
            "dependability": "Professional Dependability"}
        
        self.labels = {
            "count": "Count",
            "avoidance": "Avoidance",
            "anxiety": "Anxiety",
            "attachment": "Attachment",
            "Trust": "Trust",
            "Confidence": "Confidence",
            "Transformation": "Transformation",
            "Dominance": "Dominance",
            "Nurture": "Nurture",
            }
        
        self.numeric_idxs = {0: ['...', '...'],
                             1: ['...'],
                             2: ['...']}
        
        self.team_filter = "Before answering the questions, please select which team you are a part of.Keep this team in mind when answering the questions. I belong to the:"
        
        self.reversed_items = ["...", "...", "...",
                    "...", "...",
                    "...", "...",
                    "...",
                    "...", "...",
                    "...", "..."]
        
        self.reformed_items = ["...", "...", "...",
                  "...", "...",
                  "...", "...",
                  "...",
                  "...", "...",
                  "...", "..."]
        
        self.attachment_idxs = ["...",
                                "..."]

        self.anxiety = ["...",
            "...",
            "...",
            "...",
            "...",
            "..."]
    
        self.avoidance = ['...',
                '...',
                '...',
                '...',
                '...',
                "...",
                '...',
                "..."]

        self.org_core_idxs = ["...",
                              "..."]
        
        self.team_core_idxs = ["...",
                               "..."]
        
        self.free_form = ["...",
                          "...",
                          "...",]
        
        self.safety_idxs = ["...",
                            "..."]
        
        self.dependability_idxs = ["...",
                                   "..."]

class Dutch:
    def __init__(self):
        self.language = "Dutch"
        self.language_code = "nl"
        
        self.scales = {
            "org_core": "Kern van de organisatie",
            "team_core": "Team kern",
            "safety": "Team veiligheid",
            "dependability": "Team betrouwbaarheid"}
        
        self.labels = {
            "count": "Aantal",
            "avoidance": "Avoidance",
            "anxiety": "Anxiety",
            "attachment": "Attachment",
            "Trust": "Vertrouwen",
            "Confidence": "Assertiviteit",
            "Transformation": "Transformatie",
            "Dominance": "Dominantie",
            "Nurture": "Verzorgen",
            }
        
        self.numeric_idxs = {0: ['...', '...'],
                             1: ['...'],
                             2: ['...']}
        
        self.team_filter = "Voor het beantwoorden van de vragen, selecteer eerst bij welk team je hoort. Houd dit team in gedachten bij het beantwoorden van de vragen. Ik hoor bij team:"
        
        self.reversed_items = ["...", "...", "...",
                    "...", "...",
                    "...", "...",
                    "...",
                    "...", "...",
                    "...", "..."]
        
        self.reformed_items = ["...", "...", "...",
                  "...", "...",
                  "...", "...",
                  "...",
                  "...", "...",
                  "...", "..."]
        
        self.attachment_idxs = ["...",
                                "..."]

        self.anxiety = ["...",
            "...",
            "...",
            "...",
            "...",
            "..."]
    
        self.avoidance = ['...',
                '...',
                '...',
                '...',
                '...',
                "...",
                '...',
                "..."]

        self.org_core_idxs = ["...",
                              "..."]
        
        self.team_core_idxs = ["...",
                               "..."]
        
        self.free_form = ["...",
                          "...",
                          "...",]
        
        self.safety_idxs = ["...",
                            "..."]
        
        self.dependability_idxs = ["...",
                                   "..."]

class Danish:
    """
    Danish language class for the LYSNA application.
    """

    def __init__(self):
        self.language = "Danish"
        self.language_code = "da"
        self.scales = {
            "org_core": "Organisationens Kerne",
            "team_core": "Teamets Kerne",
            "safety": "Psykologisk Sikkerhed",
            "dependability": "Professionel Pålidelighed"}
        self.labels = {
            "count": "Antal",
            "avoidance": "Avoidance",
            "anxiety": "Anxiety",
            "attachment": "Attachment",
            "Trust": "Tillid",
            "Confidence": "Mod",
            "Transformation": "Transformation",
            "Dominance": "Dominans",
            "Nurture": "Omsorg",
            }
        
        self.numeric_idxs = {0: ['...', '...'],
                             1: ['...'],
                             2: ['...']}
        
        self.team_filter = "Inden du besvarer spørgsmålene, bedes du vælge hvilket team du er en del af. Husk dette team, når du besvarer spørgsmålene. Jeg tilhører:"
        
        self.reversed_items = ["...", "...", "...",
                    "...", "...",
                    "...", "...",
                    "...",
                    "...", "...",
                    "...", "..."]
        
        self.reformed_items = ["...", "...", "...",
                  "...", "...",
                  "...", "...",
                  "...",
                  "...", "...",
                  "...", "..."]
        
        self.attachment_idxs = ["...",
                                "..."]

        self.anxiety = ["...",
            "...",
            "...",
            "...",
            "...",
            "..."]
    
        self.avoidance = ['...',
                '...',
                '...',
                '...',
                '...',
                "...",
                '...',
                "..."]

        self.org_core_idxs = ["...",
                              "..."]
        
        self.team_core_idxs = ["...",
                               "..."]
        
        self.free_form = ["...",
                          "...",
                          "...",]
        
        self.safety_idxs = ["...",
                            "..."]
        
        self.dependability_idxs = ["...",
                                   "..."]