# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [1.0.0] - 2018-11-30

## G�n�ral
- la couleur des pions n'est plus g�rer avec .ColorIndex mais avec .Color = RGB(x,y,z)
- respect de la norme "prefixer les variables pass�s en param�tres par un p"
- d�finition ajout� pour chaque fonction
- toutes les fonctions sont d�clar�s explicitement Public ou Private

## Feuille
- ajout Cancel = true sur l'evenement BeforeDoubleClic 

## Module

### Tools
- ajout module Tools 
- proc�dure Sleep() a �t� d�plac� du module Game vers le module Tools
- proc�dure Sleep() est maintenant comment� afin de ne plus g�n�rer d'erreur 
- ajout foncion IsInArray()
- ajout foncion IsArrayNullOrEmpty()
- ajout foncion MakeBlueprintFromBoard()
- ajout proc�dure Compute()

### Game
- suppression module Game 
- proc�dure Play() a �t� d�plac� du module Game vers le module Tools
- renomme proc�dure Play() en Run()

### BoardConstructeur
- ajout proc�dure SetNameRanged pour initialiser les plages nomm�es (�tait avant effectu� dans l'instanciation d'un objet BoardModel)

## Module de classe

### PawnModel
- update proc�dure Build()
- update propri�t�s Get/Let
- update foncion TryMoveTo()
- ajout proc�dure CanMoveTo()
- ajout proc�dure priv�e MoveTo()
- ajout des variables priv�es prvColor, prvIsPawn, prvIsQueen

### MoveModel
- l'objet move ne v�rifie plus si la variable priv�e prvPawn est un v�ritable pion
- suppression de la propri�t� IsMoveOrAttack() 

### YouShallNotPassModel
- ajout module de classe YouShallNotPassModel
- ajout proc�dure Snapshot()
- ajout propri�t� IsSuccess() 