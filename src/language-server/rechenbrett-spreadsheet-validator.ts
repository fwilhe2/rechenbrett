import { ValidationAcceptor, ValidationChecks } from 'langium';
import { RechenbrettSpreadsheetAstType, Person } from './generated/ast';
import type { RechenbrettSpreadsheetServices } from './rechenbrett-spreadsheet-module';

/**
 * Register custom validation checks.
 */
export function registerValidationChecks(services: RechenbrettSpreadsheetServices) {
    const registry = services.validation.ValidationRegistry;
    const validator = services.validation.RechenbrettSpreadsheetValidator;
    const checks: ValidationChecks<RechenbrettSpreadsheetAstType> = {
        Person: validator.checkPersonStartsWithCapital
    };
    registry.register(checks, validator);
}

/**
 * Implementation of custom validations.
 */
export class RechenbrettSpreadsheetValidator {

    checkPersonStartsWithCapital(person: Person, accept: ValidationAcceptor): void {
        if (person.name) {
            const firstChar = person.name.substring(0, 1);
            if (firstChar.toUpperCase() !== firstChar) {
                accept('warning', 'Person name should start with a capital.', { node: person, property: 'name' });
            }
        }
    }

}
