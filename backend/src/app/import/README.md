# Excel recipe import

The importer uses the DTOs already defined in `src/app/entity/contracts/repository`: `IngredientInDTO`, `IngredientOutDTO`, `RecipeInDTO`, and `RecipeOutDTO`. It does not redefine these DTOs. `WorkbookRecipe` is an internal parsing structure used only before workbook ingredient names are resolved to database IDs.

The original `RecipeRepository` contract in this snapshot does not declare an `addRecipe` method. To avoid silently changing an existing project contract, `importRecipesWorkbook` receives a `saveRecipe` callback typed with the existing `RecipeInDTO` and `RecipeOutDTO`. Connect that callback to the team's recipe service/repository when its implementation is available.

Run from `backend` after installing dependencies:

```sh
npm install
npm test
npm run build
```

The source archive contains no concrete database repository or schema, so persistence cannot safely be wired to PostgreSQL here. The current import calls writes sequentially; use a database transaction in the eventual persistence adapter if all-or-nothing import is required.
