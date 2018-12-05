# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [1.1.0] - 2018-12-04
## Feuille1
- update utilisation du BotManager
- ajout vérification de bot (YouShallNotPass)
- ajout vérification d'égalité 

### BoardConstructeur
- suppression module BoardConstructeur 

### BotManager
- ajout module BotManager
- ajout procédure Run()
- ajout procédure RunBot()

### BoardModel
- procédure FormatBoard() a été déplacée du module BoardConstructeur vers le module BoardModel
- procédure Initialisation() a été déplacée du module BoardConstructeur vers le module BoardModel
- procédure SetNamedRange() a été déplacée du module BoardConstructeur vers le module BoardModel
- update procédure SearchWinner (détecte maintenant les égalité)

## [1.0.0] - 2018-12-02

## Général
- la couleur des pions n'est plus gérée avec .ColorIndex mais avec .Color = RGB(x,y,z)
- respect de la norme "préfixer les variables passées en paramètres par un 'p'"
- définition ajouté pour chaque fonction
- toutes les fonctions sont déclarées explicitement Public ou Private
- update générale => Option Explicit, toutes les variables doivent être déclarées             
- update générale => Option Base 0, les Arrays comment à "0" (Option Base 1 => les Arrays commence à 1)
                               
## Feuille1
- ajout Cancel = true sur l'évènement BeforeDoubleClic

## Module

### Enums
- ajout fonction EnumString() 

### Tools
- ajout module Tools 
- procédure Sleep() a été déplacée du module Game vers le module Tools
- procédure Sleep() est maintenant commenté afin de ne plus générer d'erreur 
- ajout foncion IsInArray()
- ajout foncion IsArrayNullOrEmpty()
- ajout foncion MakeBlueprintFromBoard()
- ajout procédure Compute()

### Game
- suppression module Game 
- procédure Play() a été déplacée du module Game vers le module Tools
- renomme procédure Play() en Run()

### BoardConstructeur
- ajout procédure SetNamedRange pour initialiser les plages nommées (était avant effectué dans l'instanciation d'un objet BoardModel)

## Module de classe

### PawnModel
- update procédure Build()
- update propriétés Get/Let
- update foncion TryMoveTo()
- ajout procédure CanMoveTo()
- ajout procédure privée MoveTo()
- ajout des variables privées prvColor, prvIsPawn, prvIsQueen

### MoveModel
- l'objet move ne vérifie plus si la variable privée prvPawn est un véritable pion
- suppression de la propriété IsMoveOrAttack() 

### YouShallNotPassModel
- ajout module de classe YouShallNotPassModel
- ajout procédure Snapshot()
- ajout propriété IsSuccess()
- ajout procédure Rollback()