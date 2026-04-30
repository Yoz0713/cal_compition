export type PaginateResult<T> = {
  items: T[];
  totalItems: number;
  totalPages: number;
  currentPage: number;
  pageSize: number;
};

export function paginate<T>(
  items: T[],
  page: number,
  pageSize: number,
): PaginateResult<T> {
  const totalItems = items.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
  const currentPage = Math.min(Math.max(1, page), totalPages);
  const start = (currentPage - 1) * pageSize;

  return {
    items: items.slice(start, start + pageSize),
    totalItems,
    totalPages,
    currentPage,
    pageSize,
  };
}
