// Interface-like abstractions for LibraryRepository
export class LibraryRepository {
  addItem(item) {
    throw new Error('Not implemented');
  }
  getItemById(id) {
    throw new Error('Not implemented');
  }
  listItems() {
    throw new Error('Not implemented');
  }
}

// In-memory implementation
export class InMemoryLibraryRepository extends LibraryRepository {
  constructor() {
    super();
    this.items = [];
  }
  addItem(item) {
    this.items.push(item);
  }
  getItemById(id) {
    return this.items.find(item => item.id === id);
  }
  listItems() {
    return this.items;
  }
}
