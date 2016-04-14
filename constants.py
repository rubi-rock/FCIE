# HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!!
#
# Same idea than the DictEnum but simpler and keep original order for the enumerated items which is useful to
# dump in a CSV file
#
class ListEnum(list):
    # Initialize the dict from another one containing the enum list Names & values, plus publish all keys as properties
    # directly available from this object
    def __init__(self, enum_list):
        # keep a trace of the list for the 'as_list' operator
        super().__init__()
        self.__enum_list = enum_list
        # load the list
        for name in enum_list:
            self.append(name)
        # create the properties
        [object.__setattr__(self, name.replace(' ', '_'), name) for name in enum_list]

    @property
    def as_list(self):
        return self.__enum_list

OUI_NON = {'T': 'Oui', 'F': 'False', None: ''}

ExcelBlockDef = ListEnum(['section', 'headers'])
ExcelBlock = ListEnum(['customer', 'user', 'md', 'institution', 'user_access'])
CSVBlocks = ListEnum(['Page2Top', 'Page2Left', 'Page2Right', 'Page3', 'Page4'])

MatchingBlocks = { CSVBlocks.Page2Top: ExcelBlock.customer, CSVBlocks.Page2Left: ExcelBlock.user, CSVBlocks.Page2Right: ExcelBlock.md,
     CSVBlocks.Page3: ExcelBlock.institution, CSVBlocks.Page4: ExcelBlock.user_access}

EXCEL_HEADERS = dict(
    customer={
        ExcelBlockDef.section: [''],
        ExcelBlockDef.headers: ['Nom du client / Agence', 'Numéro d''agence', 'Mot de passe TIP-I', 'Ville']
    },
    user={
        ExcelBlockDef.section: ['Utilisateurs'],
        ExcelBlockDef.headers: ['Nom d''utilisateur (Username)', 'Mot de passe', 'Nom', 'Prénom']
    },
    md={
        ExcelBlockDef.section: ['Médecins'],
        ExcelBlockDef.headers: ['Numéro de pratique', 'Numéro de groupe', 'Nom', 'Prénom', 'Spécialité', 'RMX (oui/non)',
                              'Inc (oui/non)', 'Nom compagnie', 'Date fin année fiscale inc']
    },
    institution={
        ExcelBlockDef.section: ['Établissements'],
        ExcelBlockDef.headers: ['Numéro de pratique', 'Numéro de groupe', 'Numéro d''établissement',
                              'Nom d''établissement',
                              'Pourcentage', 'RMX (oui/non)', 'Secteurs Cabinet', 'Secteurs CLSC',
                              'Secteurs Centre hospitalier']
    },
    user_access={
        ExcelBlockDef.section: ['Accès utilisateurs', 'Médecin-Groupe'],
        ExcelBlockDef.headers: ['Nom d''utilisateur', 'Médecin-Groupe']
    })

