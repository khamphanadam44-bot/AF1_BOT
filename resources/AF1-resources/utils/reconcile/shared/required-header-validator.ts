import { canonicalHeader } from "../../validators/shared/header-matcher";

export const assertRequiredHeaders = (
  reportCode: string,
  sourceName: string,
  actualHeaders: readonly string[],
  requiredHeaders: readonly string[],
): void => {
  const actualHeaderSet = new Set(
    actualHeaders
      .filter((header) => header.trim() !== "")
      .map(canonicalHeader),
  );
  const missingHeaders = requiredHeaders.filter(
    (header) => !actualHeaderSet.has(canonicalHeader(header)),
  );

  if (missingHeaders.length > 0) {
    throw new Error(
      `[${reportCode}] ${sourceName} missing header(s): ` +
        missingHeaders.join(", "),
    );
  }
};
