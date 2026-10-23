import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import { importRecipesWorkbook, parseRecipesWorkbook, WorkbookImportError } from './recipe-import.js';
import type { IngredientInDTO, IngredientOutDTO, IngredientRepository } from '../entity/contracts/repository/ingredient.repo.js';
import type { RecipeInDTO, RecipeOutDTO } from '../entity/contracts/repository/recipe.repo.js';

function workbookBuffer(): Buffer {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Название', 'Калории\n(ккал на 100 г)', 'Белки\n(г на 100 г)', 'Жиры\n(г на 100 г)', 'Углеводы\n(г на 100 г)', 'Проверка'],
    ['Рис', 130, 2.7, 0.3, 28, 'OK'], ['Вода', 0, 0, 0, 0, 'OK'],
  ]), 'Ингредиенты');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Название', 'Описание', 'Проверка'], ['Каша', 'Простая каша', 'OK']]), 'Рецепты');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Рецепт', 'Ингредиент', 'Количество, г', 'Своё количество\n(целое число)', 'Своя единица', 'Проверка'],
    ['Каша', 'Рис', 100, null, null, 'OK'], ['Каша', 'Вода', 200, null, null, 'OK'],
  ]), 'Состав');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
}
function repositories() {
  let nextIngredientId = 1;
  let nextRecipeId = 1;
  const ingredients: IngredientOutDTO[] = [];
  const recipes: RecipeOutDTO[] = [];
  const ingredientRepository: Pick<IngredientRepository, 'addIngredient'> = {
    async addIngredient(dto: IngredientInDTO) {
      const item: IngredientOutDTO = { id: nextIngredientId++, name: dto.name, calories: dto.calories ?? null, proteins: dto.proteins ?? null, fats: dto.fats ?? null, carbohydrates: dto.carbohydrates ?? null };
      ingredients.push(item);
      return item;
    },
  };
  const saveRecipe = async (dto: RecipeInDTO): Promise<RecipeOutDTO> => {
    const item: RecipeOutDTO = {
      id: nextRecipeId++, name: dto.name, description: dto.description,
      ingredients: dto.ingredients.map((i) => {
        const ingredient = ingredients.find((x) => x.id === i.ingredientId);
        if (!ingredient) throw Error('ingredient not found');
        return { ingredient, amountGr: i.amountGr, customAmount: i.customAmount, customUnit: i.customUnit };
      }),
    };
    recipes.push(item);
    return item;
  };
  return { ingredientRepository, saveRecipe, ingredients, recipes };
}

test('parses workbook using existing IngredientInDTO and RecipeInDTO field names', () => {
  const result = parseRecipesWorkbook(workbookBuffer());
  assert.equal(result.ingredients.length, 2);
  assert.equal(result.ingredients[0].name, 'Рис');
  assert.equal(result.recipes[0].ingredients[0].amountGr, 100);
});
test('rejects a workbook without required sheets', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Название']]), 'Ингредиенты');
  assert.throws(() => parseRecipesWorkbook(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer), WorkbookImportError);
});
test('resolves ingredient names to IDs and passes the existing RecipeInDTO to the save callback', async () => {
  const repos = repositories();
  const result = await importRecipesWorkbook(workbookBuffer(), repos.ingredientRepository, repos.saveRecipe);
  assert.deepEqual(result, { ingredientsCreated: 2, recipesCreated: 1 });
  assert.equal(repos.recipes[0].ingredients[0].ingredient.name, 'Рис');
  assert.equal(repos.recipes[0].ingredients[0].amountGr, 100);
});
test('does not write anything when workbook validation fails', async () => {
  const repos = repositories();
  await assert.rejects(() => importRecipesWorkbook(Buffer.from('not xlsx'), repos.ingredientRepository, repos.saveRecipe), WorkbookImportError);
  assert.equal(repos.ingredients.length, 0);
  assert.equal(repos.recipes.length, 0);
});
