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

OUI_NON = {True: 'Oui', False: 'Non', None: 'Non'}
BOOL_OUI_NON = {True: 'Oui', False: ' ', None: ' '}

ExcelBlock = ListEnum(['attention_de', 'agent_purkinje', 'proposition', 'site', 'customer', 'user', 'md', 'institution', 'user_access'])
ExcelBlockDef = ListEnum(['section', 'headers'])

EXCEL_HEADERS = dict(
    attention_de={
        ExcelBlockDef.section: [],
        ExcelBlockDef.headers: ['Nom du client / Agence:', 'Attention', 'Rue', 'Ville', 'Code postal']
    },
    agent_purkinje={
        ExcelBlockDef.section: [],
        ExcelBlockDef.headers: ['Date:', 'Agent Purkinje:', 'Téléphone:', 'Fax:', 'Courriel']
    },
    proposition={
        ExcelBlockDef.section: [],
        ExcelBlockDef.headers: ['Numéro de pratique', 'Nom', 'Prénom', 'Spécialité', 'Qté', 'Mensualité', 'Mensualité\nTotale']
    },
    site={
        ExcelBlockDef.section: [],
        ExcelBlockDef.headers: ['Nom', 'Adresse,Ville', 'Code Postal', 'Province', 'Pays', 'Telephone', 'Fax']
    },
    customer={
        ExcelBlockDef.section: [],
        ExcelBlockDef.headers: ['Nom du client / Agence', 'Numéro d''agence', 'Mot de passe TIP-I', 'Ville']
    },
    user={
        ExcelBlockDef.section: ['Utilisateurs'],
        ExcelBlockDef.headers: ['Nom d''utilisateur (Username)', 'Mot de passe', 'Nom', 'Prénom', 'System d''opération']
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
        ExcelBlockDef.section: ['Accès utilisateurs'],
        ExcelBlockDef.headers: ['Nom d''utilisateur', 'Médecin-Groupe']
    })


COL_SIZE = dict(
    #                   A   B  C   D   E   F   G   H   I
    proposition     = [25, 25, 25, 25, 25, 25, 25, 25, 25],
    customer        = [20, 20, 20, 20, 60, 15, 15, 15, 20],
    institution     = [15, 15, 15, 15, 15, 15, 15, 15, 25],
    user_access     = [25, 25]
)

OS_LIST = ['Windows XP', 'Windows 7, 8 ou 10', 'Windows 2008 ou 20012', 'OSX (Apple)']