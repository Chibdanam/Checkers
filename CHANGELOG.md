# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [1.1.0] - 2018-12-04
## Feuille1
- update utilisation du BotManager
- ajout v�rification de bot (YouShallNotPass)
- ajout v�rification d'�galit� 

### BoardConstructeur
- suppression module BoardConstructeur 

### BotManager
- ajout module BotManager
- ajout proc�dure Run()
- ajout proc�dure RunBot()

### BoardModel
- proc�dure FormatBoard() a �t� d�plac�e du module BoardConstructeur vers le module BoardModel
- proc�dure Initialisation() a �t� d�plac�e du module BoardConstructeur vers le module BoardModel
- proc�dure SetNamedRange() a �t� d�plac�e du module BoardConstructeur vers le module BoardModel
- update proc�dure SearchWinner (d�tecte maintenant les �galit�)

## [1.0.0] - 2018-12-02

## G�n�ral
- la couleur des pions n'est plus g�r�e avec .ColorIndex mais avec .Color = RGB(x,y,z)
- respect de la norme "pr�fixer les variables pass�es en param�tres par un 'p'"
- d�finition ajout� pour chaque fonction
- toutes les fonctions sont d�clar�es explicitement Public ou Private
- update g�n�rale => Option Explicit, toutes les variables doivent �tre d�clar�es             
- update g�n�rale => Option Base 0, les Arrays comment � "0" (Option Base 1 => les Arrays commence � 1)
                               
## Feuille1
- ajout Cancel = true sur l'�v�nement BeforeDoubleClic

## Module

### Enums
- ajout fonction EnumString() 

### Tools
- ajout module Tools 
- proc�dure Sleep() a �t� d�plac�e du module Game vers le module Tools
- proc�dure Sleep() est maintenant comment� afin de ne plus g�n�rer d'erreur 
- ajout foncion IsInArray()
- ajout foncion IsArrayNullOrEmpty()
- ajout foncion MakeBlueprintFromBoard()
- ajout proc�dure Compute()

### Game
- suppression module Game 
- proc�dure Play() a �t� d�plac�e du module Game vers le module Tools
- renomme proc�dure Play() en Run()

### BoardConstructeur
- ajout proc�dure SetNamedRange pour initialiser les plages nomm�es (�tait avant effectu� dans l'instanciation d'un objet BoardModel)

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
- ajout proc�dure Rollback()