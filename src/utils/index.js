export async function errorCaptured(asyncFunc) {
  try {
    const res = await asyncFunc;
    return [null, res];
  } catch (error) {
    return [error, null]
  }
}
