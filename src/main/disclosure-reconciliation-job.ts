type RunInput = { targetDate: string; idempotencyKey: string };

export function createDisclosureReconciliationJob(dependencies: {
  targetDate(): string;
  disclosure: { run(input: RunInput): Promise<unknown> };
  reconciliation: { run(input: RunInput): Promise<unknown> };
}) {
  return async (idempotencyKey: string): Promise<unknown> => {
    const targetDate = dependencies.targetDate();
    const result = await dependencies.disclosure.run({ targetDate, idempotencyKey });
    await dependencies.reconciliation.run({
      targetDate,
      idempotencyKey: `material-reconciliation:${idempotencyKey}`,
    });
    return result;
  };
}
