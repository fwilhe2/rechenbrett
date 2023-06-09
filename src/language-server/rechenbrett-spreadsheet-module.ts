import {
    createDefaultModule, createDefaultSharedModule, DefaultSharedModuleContext, inject,
    LangiumServices, LangiumSharedServices, Module, PartialLangiumServices
} from 'langium';
import { RechenbrettSpreadsheetGeneratedModule, RechenbrettSpreadsheetGeneratedSharedModule } from './generated/module';
import { RechenbrettSpreadsheetValidator, registerValidationChecks } from './rechenbrett-spreadsheet-validator';

/**
 * Declaration of custom services - add your own service classes here.
 */
export type RechenbrettSpreadsheetAddedServices = {
    validation: {
        RechenbrettSpreadsheetValidator: RechenbrettSpreadsheetValidator
    }
}

/**
 * Union of Langium default services and your custom services - use this as constructor parameter
 * of custom service classes.
 */
export type RechenbrettSpreadsheetServices = LangiumServices & RechenbrettSpreadsheetAddedServices

/**
 * Dependency injection module that overrides Langium default services and contributes the
 * declared custom services. The Langium defaults can be partially specified to override only
 * selected services, while the custom services must be fully specified.
 */
export const RechenbrettSpreadsheetModule: Module<RechenbrettSpreadsheetServices, PartialLangiumServices & RechenbrettSpreadsheetAddedServices> = {
    validation: {
        RechenbrettSpreadsheetValidator: () => new RechenbrettSpreadsheetValidator()
    }
};

/**
 * Create the full set of services required by Langium.
 *
 * First inject the shared services by merging two modules:
 *  - Langium default shared services
 *  - Services generated by langium-cli
 *
 * Then inject the language-specific services by merging three modules:
 *  - Langium default language-specific services
 *  - Services generated by langium-cli
 *  - Services specified in this file
 *
 * @param context Optional module context with the LSP connection
 * @returns An object wrapping the shared services and the language-specific services
 */
export function createRechenbrettSpreadsheetServices(context: DefaultSharedModuleContext): {
    shared: LangiumSharedServices,
    RechenbrettSpreadsheet: RechenbrettSpreadsheetServices
} {
    const shared = inject(
        createDefaultSharedModule(context),
        RechenbrettSpreadsheetGeneratedSharedModule
    );
    const RechenbrettSpreadsheet = inject(
        createDefaultModule({ shared }),
        RechenbrettSpreadsheetGeneratedModule,
        RechenbrettSpreadsheetModule
    );
    shared.ServiceRegistry.register(RechenbrettSpreadsheet);
    registerValidationChecks(RechenbrettSpreadsheet);
    return { shared, RechenbrettSpreadsheet };
}
