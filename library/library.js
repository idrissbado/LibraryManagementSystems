// Main Library class using dependency injection
import { Observable } from './observer.js';
import { LibraryItemFactory } from './factory.js';

export class Library extends Observable {
  constructor(repository, lendingStrategy) {
    super();
    this.repository = repository;
    this.lendingStrategy = lendingStrategy;
  }

  addItem(type, props) {
    const item = LibraryItemFactory.createItem(type, props);
    this.repository.addItem(item);
    this.notify({ event: 'itemAdded', item });
  }

  lendItem(user, itemId) {
    const item = this.repository.getItemById(itemId);
    if (!item) throw new Error('Item not found');
    if (this.lendingStrategy.canLend(user, item)) {
      this.notify({ event: 'itemLent', item, user });
      return true;
    }
    return false;
  }

  listItems() {
    return this.repository.listItems();
  }
}
