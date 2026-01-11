// Strategy pattern for different lending policies
export class LendingStrategy {
  canLend(user, item) {
    throw new Error('Not implemented');
  }
}

export class StandardLendingStrategy extends LendingStrategy {
  canLend(user, item) {
    return !user.hasOverdueItems;
  }
}

export class PremiumLendingStrategy extends LendingStrategy {
  canLend(user, item) {
    return true; // Premium users can always borrow
  }
}
