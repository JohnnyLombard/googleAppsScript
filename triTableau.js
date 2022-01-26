// Exemple de function pour trier un tableau

let reservedCars = [];

let car1 = [240, "Voiture 1"];
let car2 = [240, "Voiture 2"];
let car3 = [200, "Voiture 3"];
let car4 = [350, "Voiture 4"];

reservedCars.push(car1, car2, car3, car4);

console.log(reservedCars);

// Maintenant on veut trier le tableau par ordre croissant de kilométrage
// Pour comprendre la fonction sort() => https://www.tutorialspoint.com/how-to-define-custom-sort-function-in-javascript

reservedCars.sort((first, second) => {
    if (first[0] > second[0]) {
        return 1;
    }
    if (first[0] < second[0]) {
        return -1;
    }
    return 0;
})

console.log(reservedCars);
