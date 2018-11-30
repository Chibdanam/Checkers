# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [1.0.0] - 2018-11-30

## Général
- la couleur des pions n'est plus gérer avec .ColorIndex mais avec .Color = RGB(x,y,z)
- respect de la norme "prefixer les variables passés en paramètres par un p"
- définition ajouté pour chaque fonction
- toutes les fonctions sont déclarés explicitement Public ou Private

## Feuille
- ajout Cancel = true sur l'evenement BeforeDoubleClic 

## Module

### Tools
- ajout module Tools 
- procédure Sleep() a été déplacé du module Game vers le module Tools
- procédure Sleep() est maintenant commenté afin de ne plus générer d'erreur 
- ajout foncion IsInArray()
- ajout foncion IsArrayNullOrEmpty()
- ajout foncion MakeBlueprintFromBoard()
- ajout procédure Compute()

### Game
- suppression module Game 
- procédure Play() a été déplacé du module Game vers le module Tools
- renomme procédure Play() en Run()

### BoardConstructeur
- ajout procédure SetNameRanged pour initialiser les plages nommées (était avant effectué dans l'instanciation d'un objet BoardModel)

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