# Générateur de cours

Une application web simple et rapide qui transforme un cours brut en un document Word propre, structuré et prêt à être révisé.

## 💡 Pourquoi ce projet ?

Tout est parti d'un constat tout bête, à la maison : ma copine, étudiante en droit, passait un temps fou à mettre en page ses cours sous Word. Titres à hiérarchiser, articles de loi à mettre en valeur, arrêts à repérer, couleurs et polices à harmoniser... Autant d'heures perdues à faire du formatage au lieu de réviser.

De mon côté, étudiant en informatique, je me suis dit qu'on pouvait automatiser tout ça. L'idée : une plateforme où l'on colle son cours tel quel, et qui le remet au propre toute seule, sans perdre de temps sur la mise en page.

Ce projet est donc né de ce besoin concret, à deux : elle apporte le regard de l'utilisatrice (ce qui compte vraiment dans un cours de droit), moi la partie technique. C'est un side project perso, qu'on améliore au fil de nos envies et de ses retours.

## 🚀 Ce que ça fait

Il suffit de coller votre texte dans l'interface, et l'application s'occupe automatiquement :

- De la hiérarchie des titres (Thèmes, Chapitres, Sections, Sous-sections, etc.)
- De la mise en valeur intelligente des références légales (articles de loi, jurisprudence, arrêts)
- De l'application d'un design soigné (palettes de couleurs au choix, polices, en-têtes, encadrés de légende)
- Du téléchargement direct d'un fichier `.docx` prêt à être imprimé ou partagé

## 🛠️ Comment l'utiliser ?

1. Ouvrez le fichier `index.html` dans un navigateur web.
2. Personnalisez vos styles ou choisissez une palette de couleurs dans le panneau de gauche.
3. Collez votre cours brut dans la zone de texte à droite.
4. Cliquez sur **Générer et télécharger le cours (.docx)** !

## 📂 Structure du projet

- `index.html` : l'interface utilisateur de l'application
- `style.css` : la feuille de style de l'interface web
- `script.js` : la logique de parsing du texte et de génération du document Word (via la librairie `docx`)