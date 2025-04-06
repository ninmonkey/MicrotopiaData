## About

Describes the core files that are generated and which source files they come from. 
( for: https://github.com/ninmonkey/MicrotopiaData/issues/18 )

## Raw files

All of the `*-raw.xlsx` files are a direct save of the original `*.fods` files. 
The styles don't fully export, but the raw data was fully preserved.

The data is cleaned up, then exported to `json`, `csv`, `xlsx` files. The `xlsx` are easier for humans to view, but are not required. 

## Checklist

- [x] biome-raw.xlsx
  - [x] Biome_Objects : `biome-objects.json`, `biome-objects-expanded.json`
  - [x] Plants : `biome-plants.json`, `biome-plants-column-desc.json`
- [x] Changelog-raw.xlsx : `changelog.csv`
- [ ] Instinct-raw.xlsx
  - [ ] Instinct
  - [ ] Instinct_Order
  - [ ] Goals_old
  - [ ] Goals_old2
- [x] loc-raw.xlsx
  - [x] Legend
  - [x] UI : `loc-ui.json`
  - [x] Objects : `loc-objects.json`
  - [x] Tutorial
  - [x] Instinct
  - [x] TechTree
  - [x] Credits
  - [x] Achievements
  - [ ] wip
- [x] prefabs-raw.xlsx
  - [x] Buildings : `prefabs-buildings.json`
  - [x] Factory_Recipes : `prefabs-factoryrecipes.json`
  - [x] Ant_Castes : `prefabs-antcastes.json`
  - [x] Pickups : `prefabs-pickups.json`
  - [x] Trails
  - [x] Pickup Categories
  - [x] Status Effects
  - [x] Hunger
- [x] sequences-raw.xlsx
  - [x] Tutorial
  - [x] Tutorial_old
  - [x] Events
- [ ] techtree-raw.xlsx
  - [ ]  Tech Tree : `techtree-techtree.json`
  - [ ] Research_Recipes
  - [ ] Exploration_Order
- others
  - [x] `crusher-output.json`