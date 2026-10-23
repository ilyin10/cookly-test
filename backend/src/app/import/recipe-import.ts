import * as XLSX from 'xlsx';
import type { IngredientInDTO, IngredientOutDTO, IngredientRepository } from '../entity/contracts/repository/ingredient.repo.js';
import type { RecipeInDTO, RecipeOutDTO } from '../entity/contracts/repository/recipe.repo.js';

/** A workbook row before ingredient names are resolved to database IDs. */
type WorkbookRecipe = {
  name: string;
  description: string;
  ingredients: Array<{
    ingredientName: string;
    amountGr: number;
    customAmount: number | null;
    customUnit: string | null;
  }>;
};

type ParsedWorkbook = { ingredients: IngredientInDTO[]; recipes: WorkbookRecipe[] };

export class WorkbookImportError extends Error {
  constructor(message: string) { super(message); this.name = 'WorkbookImportError'; }
}

const REQUIRED_SHEETS = ['Ингредиенты', 'Рецепты', 'Состав'] as const;
const normalize = (value: string) => value.trim().toLocaleLowerCase('ru');
const text = (value: unknown): string => value == null ? '' : String(value).trim();

function nullableNumber(value: unknown, label: string, rowNumber: number): number | null {
  if (value == null || text(value) === '') return null;
  const number = typeof value === 'number' ? value : Number(text(value).replace(',', '.'));
  if (!Number.isFinite(number) || number < 0) {
    throw new WorkbookImportError(`${label}: ожидается неотрицательное число в строке ${rowNumber}`);
  }
  return number;
}
function rowsOf(workbook: XLSX.WorkBook, name: string): unknown[][] {
  const sheet = workbook.Sheets[name];
  if (!sheet) throw new WorkbookImportError(`В файле отсутствует лист «${name}»`);
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true });
  if (!rows.length) throw new WorkbookImportError(`Лист «${name}» пуст`);
  return rows;
}
function assertHeaders(rows: unknown[][], expected: string[], sheet: string): void {
  const actual = rows[0].map(text);
  expected.forEach((header, i) => {
    if (actual[i] !== header) throw new WorkbookImportError(
      `Лист «${sheet}»: столбец ${i + 1} должен называться «${header}», получено «${actual[i] ?? ''}»`,
    );
  });
}
function dataRows(rows: unknown[][]) {
  return rows.slice(1).map((values, index) => ({ values, rowNumber: index + 2 }))
    .filter(({ values }) => values.some((value) => text(value) !== ''));
}

export function parseRecipesWorkbook(input: Buffer | Uint8Array): ParsedWorkbook {
  let workbook: XLSX.WorkBook;
  try { workbook = XLSX.read(input, { type: 'buffer' }); }
  catch { throw new WorkbookImportError('Не удалось прочитать файл .xlsx'); }
  for (const name of REQUIRED_SHEETS) if (!workbook.Sheets[name]) {
    throw new WorkbookImportError(`В файле отсутствует лист «${name}»`);
  }
  const ingredientRows = rowsOf(workbook, 'Ингредиенты');
  const recipeRows = rowsOf(workbook, 'Рецепты');
  const compositionRows = rowsOf(workbook, 'Состав');
  assertHeaders(ingredientRows, ['Название', 'Калории\n(ккал на 100 г)', 'Белки\n(г на 100 г)', 'Жиры\n(г на 100 г)', 'Углеводы\n(г на 100 г)'], 'Ингредиенты');
  assertHeaders(recipeRows, ['Название', 'Описание'], 'Рецепты');
  assertHeaders(compositionRows, ['Рецепт', 'Ингредиент', 'Количество, г', 'Своё количество\n(целое число)', 'Своя единица'], 'Состав');

  const ingredients: IngredientInDTO[] = [];
  const ingredientNames = new Set<string>();
  for (const { values, rowNumber } of dataRows(ingredientRows)) {
    const name = text(values[0]);
    if (!name) throw new WorkbookImportError(`Ингредиенты: не заполнено название в строке ${rowNumber}`);
    const key = normalize(name);
    if (ingredientNames.has(key)) throw new WorkbookImportError(`Ингредиенты: повторяется название «${name}» (строка ${rowNumber})`);
    ingredientNames.add(key);
    ingredients.push({ name,
      calories: nullableNumber(values[1], 'Калории', rowNumber),
      proteins: nullableNumber(values[2], 'Белки', rowNumber),
      fats: nullableNumber(values[3], 'Жиры', rowNumber),
      carbohydrates: nullableNumber(values[4], 'Углеводы', rowNumber),
    });
  }

  const recipesByName = new Map<string, WorkbookRecipe>();
  for (const { values, rowNumber } of dataRows(recipeRows)) {
    const name = text(values[0]); const description = text(values[1]);
    if (!name || !description) throw new WorkbookImportError(`Рецепты: требуется название и описание (строка ${rowNumber})`);
    const key = normalize(name);
    if (recipesByName.has(key)) throw new WorkbookImportError(`Рецепты: повторяется название «${name}»`);
    recipesByName.set(key, { name, description, ingredients: [] });
  }

  const compositions = new Set<string>();
  for (const { values, rowNumber } of dataRows(compositionRows)) {
    const recipeName = text(values[0]); const ingredientName = text(values[1]);
    if (!recipeName || !ingredientName) throw new WorkbookImportError(`Состав: укажите рецепт и ингредиент в строке ${rowNumber}`);
    const recipe = recipesByName.get(normalize(recipeName));
    if (!recipe) throw new WorkbookImportError(`Рецепт «${recipeName}» из строки ${rowNumber} не найден на листе «Рецепты»`);
    if (!ingredientNames.has(normalize(ingredientName))) throw new WorkbookImportError(`Ингредиент «${ingredientName}» из строки ${rowNumber} не найден на листе «Ингредиенты»`);
    const key = `${normalize(recipeName)}\0${normalize(ingredientName)}`;
    if (compositions.has(key)) throw new WorkbookImportError(`Ингредиент «${ingredientName}» повторяется в рецепте «${recipeName}»`);
    compositions.add(key);
    const amountGr = nullableNumber(values[2], 'Количество, г', rowNumber);
    const customAmount = nullableNumber(values[3], 'Своё количество', rowNumber);
    const customUnit = text(values[4]) || null;
    if ((amountGr == null || amountGr === 0) && (customAmount == null || !customUnit)) {
      throw new WorkbookImportError(`Состав: в строке ${rowNumber} укажите количество в граммах или своё количество вместе с единицей`);
    }
    recipe.ingredients.push({ ingredientName, amountGr: amountGr ?? 0, customAmount, customUnit });
  }
  if (!ingredients.length) throw new WorkbookImportError('На листе «Ингредиенты» нет данных');
  if (!recipesByName.size) throw new WorkbookImportError('На листе «Рецепты» нет данных');
  for (const recipe of recipesByName.values()) if (!recipe.ingredients.length) {
    throw new WorkbookImportError(`У рецепта «${recipe.name}» нет состава`);
  }
  return { ingredients, recipes: [...recipesByName.values()] };
}

export type RecipeImportResult = { ingredientsCreated: number; recipesCreated: number };

/** Uses the repository DTOs already defined by the backend; workbook-only names are resolved to IngredientOutDTO IDs. */
/**
 * Imports validated workbook data using the existing DTO contracts.
 * The current main branch's RecipeRepository does not expose addRecipe, so recipe persistence
 * is injected as a callback instead of pretending that such a method already exists.
 */
export async function importRecipesWorkbook(
  input: Buffer | Uint8Array,
  ingredientRepository: Pick<IngredientRepository, 'addIngredient'>,
  saveRecipe: (inDTO: RecipeInDTO) => Promise<RecipeOutDTO>,
): Promise<RecipeImportResult> {
  const parsed = parseRecipesWorkbook(input); // Validate entire workbook before making writes.
  const ingredientsByName = new Map<string, IngredientOutDTO>();
  for (const ingredient of parsed.ingredients) {
    const created = await ingredientRepository.addIngredient(ingredient);
    ingredientsByName.set(normalize(created.name), created);
  }
  for (const recipe of parsed.recipes) {
    const dto: RecipeInDTO = {
      name: recipe.name,
      description: recipe.description,
      ingredients: recipe.ingredients.map((item) => {
        const ingredient = ingredientsByName.get(normalize(item.ingredientName));
        if (!ingredient) throw new WorkbookImportError(`Не удалось получить ID ингредиента «${item.ingredientName}»`);
        return { ingredientId: ingredient.id, amountGr: item.amountGr, customAmount: item.customAmount, customUnit: item.customUnit };
      }),
    };
    await saveRecipe(dto);
  }
  return { ingredientsCreated: parsed.ingredients.length, recipesCreated: parsed.recipes.length };
}
