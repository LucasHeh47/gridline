// How VBA code runs. The interpreter is written as generators: they yield a promise whenever the
// code has to wait (a message box, Application.Wait, a pause that keeps the app responsive). The
// same code then runs two ways:
//  • macros: asynchronously, waiting for each promise;
//  • functions used in cells (=MYFUNCTION(A1)): synchronously, inside a recalculation, where
//    waiting isn't possible, so anything that would wait is reported as an error instead.
import { VbError } from './values';

/** Interpreter work: yields promises it must wait for and finally returns a T. */
export type Gen<T> = Generator<Promise<unknown>, T, unknown>;

/** A result that may be ready, may need waiting for, or may still need running (a class method). */
export type Pending<T> = T | Promise<T> | Gen<T>;

export function isGen(v: unknown): v is Gen<unknown> {
  return Object.prototype.toString.call(v) === '[object Generator]';
}

/** Turns a maybe-pending result into interpreter work. */
export function* settle<T>(v: Pending<T>): Gen<T> {
  if (v instanceof Promise) return (yield v) as T;
  if (isGen(v)) return (yield* v) as T;
  return v;
}

/** Runs interpreter work, waiting whenever it asks to (macros). */
export async function drive<T>(g: Gen<T>): Promise<T> {
  let r = g.next();
  while (!r.done) {
    let input: unknown;
    let failed = false;
    let err: unknown;
    try {
      input = await r.value;
    } catch (e) {
      failed = true;
      err = e;
    }
    r = failed ? g.throw(err) : g.next(input);
  }
  return r.value;
}

/**
 * Runs interpreter work to the end without waiting (functions in cells). Code that tries to wait
 * gets an error it can handle with On Error, like any other.
 */
export function driveSync<T>(g: Gen<T>): T {
  let r = g.next();
  while (!r.done) {
    // Nothing waits for it, so make sure a rejection isn't reported as unhandled
    r.value.catch(() => undefined);
    r = g.throw(new VbError(1004, "A function used in a cell can't wait or show messages while the sheet calculates."));
  }
  return r.value;
}
