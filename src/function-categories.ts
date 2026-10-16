// The categories of Excel's Insert Function window in its order, as xlcalc's FUNCTION_CATEGORIES gives them: the main entry keeps
// its own copy, xlcalc's root carrying the engine (a test checks the copy against the root).
import type { FunctionCategory } from 'xlcalc/contract'

export const CATEGORY_ORDER: readonly FunctionCategory[] = ['financial', 'date', 'math', 'statistical', 'lookup', 'database', 'text', 'logical', 'information', 'engineering', 'compatibility']
