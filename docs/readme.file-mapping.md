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
- [ ] loc-raw.xlsx
  - [x] Legend
  - [ ] UI : `loc-ui.json`
  - [ ] Objects : `loc-objects.json`
  - [ ] Tutorial
  - [ ] Instinct
  - [ ] TechTree
  - [ ] Credits
  - [ ] Achievements
  - [ ] wip
- [ ] prefabs-raw.xlsx
  - [ ] Buildings : `prefabs-buildings.json`
  - [ ] Factory_Recipes : `prefabs-factoryrecipes.json`
  - [ ] Ant_Castes : `prefabs-antcastes.json`
  - [ ] Pickups : `prefabs-pickups.json`
  - [ ] Trails
  - [ ] Pickup Categories
  - [ ] Status Effects
  - [ ] Hunger
- [ ] sequences-raw.xlsx
  - [ ] Tutorial
  - [ ] Tutorial_old
  - [ ] Events
- [ ] techtree-raw.xlsx
  - [ ]  Tech Tree : `techtree-techtree.json`
  - [ ] Research_Recipes
  - [ ] Exploration_Order
- others
  - [x] `crusher-output.json`