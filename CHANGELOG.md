# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [1.0.0] - 2018-11-25

## Général
- La couleur des pions n'est plus gérer avec .ColorIndex mais avec .Color = RGB(x,y,z)

## Feuille
- ajout Cancel = true sur l'evenement BeforeDoubleClic

## Module

### Tools
- Le module Tools a été ajouté
- La méthode IsInArray a été ajouté au module Tools
- La méthode IsArrayNullOrEmpty a été ajouté au module Tools
- La méthode Sleep a été déplacé  du module Game vers le module Tools
- La méthode Sleep est maintenant commenté afin de ne plus générer d'erreur 

### Game
- Le module Game est renommé Player
- La procédure Play du module Player est renommé Run

### BoardConstructeur
- Ajout methode SetNameRanged pour initiliser les plages nommées (était avant effectué dans l'instanciation d'un objet BoardModel)

## Module de classe

### PawnModel
- update procédure Build(), tryMoveTo()
- ajout des variables privées prvColor, prvIsPawn, prvIsQueen
- ajout propriété CanMoveTo()
- ajout procédure privée MoveTo()

### MoveModel
- update général l'objet move ne vérifie plus si la variable privé prvPawn est un véritable pion 
