package com.javapapers.jee;

public class AnimalTypeService {
	public String animalType(String animal) {
		String animalType = "";
		if ("Lion".equals(animal)) {
			animalType = "Wild";
		} else if ("Dog".equals(animal)) {
			animalType = "Domestic";
		} else if ("Jaguar".equals(animal) || "jaguar".equals(animal)) {
				animalType = "South American cat";
		}
		else if ("Alligator".equals(animal) || "alligator".equals(animal)) {
			animalType = "SouthEastern US animal";
		}else {
			animalType = "I don't know!";
		}
		return animalType;
	}
}
