# Instructions pour Claude Code - Projet ARES MVBA

## Contexte du projet
ARES est un projet MicroStation VBA open source, qui offre des outils d'optimisation pour la réalisations de plans 2D. 

## Standards de codage
- **Déclarations** : `Option Explicit` en première ligne obligatoire
- **Fonctions** : CamelCase (ex: `ProcessSelectedElements`)
- **Variables** : camelCase avec préfixe type (ex: `strFileName`, `objElement`)
- **Constantes** : UPPER_CASE (ex: `MAX_ELEMENTS`)
- **Gestion d'erreurs** : Toujours utiliser `On Error GoTo ErrorHandler` avec ErrorHandlerClass

## Compatibilité obligatoire
- **Version cible** : MicroStation CONNECT Edition (64 bits)
- **APIs Windows** : Déclarations 64 bits uniquement
- **Types de données** : LongPtr au lieu de Long pour les handles

## Documentation technique disponible

### Structure de la documentation extraite
```
docs/extracted_chm/
├── MicroStationVBA.hhc          # 📋 Table des matières (2 181 Ko) - CONSULTER EN PREMIER
├── MicroStationVBA.hhk          # 🔍 Index alphabétique (428 Ko)
├── html/                        # 📄 Documentation API détaillée
└── vbaconcept/html/
    ├── usvba_common_mistakes.htm      # ⚠️ PRIORITÉ ABSOLUE - Erreurs courantes
    ├── usvba_64bit_processes.htm      # 🔧 Compatibilité CONNECT Edition 64 bits
    └── usvba_calling_dll_functions.htm # 🔗 Intégration DLL/MDL
```

### Ordre de consultation obligatoire pour Claude
1. **TOUJOURS commencer par** `usvba_common_mistakes.htm` - Éviter les pièges
2. **Vérifier** `usvba_64bit_processes.htm` - Assurer compatibilité CONNECT Edition
3. **Consulter** `MicroStationVBA.hhc` - Navigation dans l'API
4. **Détailler avec** `html/` - Spécifications précises des méthodes
5. **Si nécessaire** `usvba_calling_dll_functions.htm` - Intégrations avancées

## Langues
- **Communication avec utilisateur** : Français exclusivement
- **Code et commentaires** : Anglais exclusivement
- **Documentation générée** : Anglais pour les explications, anglais pour le code

### Exemple de fonction conforme :
```vba
' Function to zoom on an element in a specified view
Public Function ZoomEl(ByVal El As element, Optional Factor As Single = 1.3) As Boolean
    On Error GoTo ErrorHandler
    Dim Rng As Range3d
    Dim PntZoom As Point3d
    Dim oView As View
    Dim pntCenter As Point3d
    Dim Pnt As Point3d

    ZoomEl = False

    ' Check if the element is graphical
    If El.IsGraphical Then
        ' Get the Last View
        Set oView = CommandState.LastView
        ' Get the range of the element
        Rng = El.Range
        ' Calculate the zoom point based on the range of the element
        With Rng
            PntZoom.X = .High.X - .Low.X
            PntZoom.Y = .High.Y - .Low.Y
            PntZoom.Z = .High.Z - .Low.Z
        End With
        ' Set the point for zooming
        With Pnt
            .X = PntZoom.X * Factor
            .Y = PntZoom.Y * Factor
            .Z = PntZoom.Z * Factor
        End With
        ' Set the view area and zoom
        oView.SetArea Rng.Low, Pnt, oView.Rotation, Rng.High.Z
        oView.ZoomAboutPoint Point3dAddScaled(Rng.Low, PntZoom, 0.5), 1
        oView.Redraw
        ZoomEl = True
    End If

    Exit Function

ErrorHandler:
    ' Return False in case of an error
    ZoomEl = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "MSGraphicalInteraction.ZoomEl"
End Function
```

## Structure du projet ARES
- `/Modules/` : Modules VBA principaux (.bas)
- `/Forms/` : Formulaires utilisateur (.frm)
- `/Classes/` : Classes personnalisées (.cls)
- `/Utils/` : Utilitaires et helpers communs
- `/Tests/` : Procédures de test et validation

### Object Model principal
```
Application
├── ActiveDesignFile        # Fichier actif
├── ActiveModelReference    # Modèle actif  
├── CommandTable           # Commandes personnalisées
└── MessageCenter          # Interface messages

ActiveModelReference
├── GetElements()          # Tous les éléments
├── GetSelectedElements()  # Éléments sélectionnés
└── AddElement()           # Ajout d'éléments

Element (classe de base)
├── Delete()               # Suppression
├── Transform()            # Transformation géométrique
├── Copy()                 # Duplication
└── Properties             # Accès propriétés
```

## Priorités de développement ARES

### Robustesse (priorité 1)
- Gestion d'erreurs exhaustive dans toutes les fonctions
- Validation des paramètres d'entrée
- Nettoyage approprié des objets COM
- Messages d'erreur informatifs pour l'utilisateur

### Performance (priorité 2)  
- Optimisation pour modèles avec milliers d'éléments
- Utilisation efficace des énumérateurs
- DoEvents pour éviter le blocage de l'interface
- Mesure et logging des performances

### Compatibilité (priorité 3)
- Code 64 bits uniquement (CONNECT Edition)
- Tests sur différentes versions de MicroStation
- Gestion gracieuse des APIs obsolètes

### Documentation (priorité 4)
- Commentaires détaillés en anglais
- Documentation utilisateur en français
- Exemples d'utilisation pour chaque fonction publique

## Instructions spéciales pour Claude

### Avant tout développement
1. **OBLIGATOIRE** : Consulter `usvba_common_mistakes.htm` pour éviter les erreurs
2. **VÉRIFIER** : Compatibilité 64 bits dans `usvba_64bit_processes.htm`
3. **RÉFÉRENCER** : API exacte dans `MicroStationVBA.hhc` et `html/`

### Qualité attendue
- **Code prêt pour production** : pas de prototypes ou d'exemples simplifiés
- **Gestion d'erreurs complète** : jamais de fonction sans On Error GoTo
- **Performance optimisée** : considérer les gros volumes de données
- **Documentation intégrée** : commentaires détaillés pour maintenance future

## Notes importantes

### Limitations connues
- MicroStation VBA ne supporte pas les collections .NET
- Attention aux références circulaires avec les objets COM
- Performance limitée pour traitement de très gros volumes (>100k éléments)

### Intégration avec MicroStation
- Respecter les conventions de nommage MicroStation
- Utiliser MessageCenter pour les notifications utilisateur
- Intégrer proprement avec le système de commandes
- Gérer correctement les états d'annulation (Undo/Redo)