import * as path from "path";
import * as fs from "fs";
import * as getNextNpmPackageVersion from "./getNextNpmPackageVersion";
import * as standardFile from "./standardFile";
import {executeCommand} from "./executeCommand"

interface AdditionalInfoError extends Error {
    additionalInfo: string;
}

function isPublishOverPreviouslyPublishVersionErrorString(errorText: string): boolean {
    const phraseMatchIndex = [
        "npm ERR! You cannot publish over the previously published version",
        "npm ERR! Cannot publish over previously published version"
    ]
        .map(phrase => phrase.toLowerCase())
        .findIndex(phrase => errorText.toLowerCase().indexOf(phrase) >= 0);
    return phraseMatchIndex >= 0;
}



function updatePackageVersion(packageJsonPath: string, version: string) {
    type Package = {version: string};
    
    const packageData: Package = standardFile.readFileJson<Package>(packageJsonPath);

    packageData.version = version;

    standardFile.writeFileJson(packageJsonPath, packageData);
}


export function deployNpmPackage(packageDirectory: string, packageName: string, packageTag: string | undefined, npmAuthToken: string): void {

    const packageJsonPath = path.join(packageDirectory, "package.json");
    const npmrcPath = path.join(packageDirectory, ".npmrc");

    console.log("Write .npmrc Deployment Token:")
    fs.writeFileSync(npmrcPath, `//registry.npmjs.org/:_authToken=${npmAuthToken}`);
    
    const maxPublishAttempts: number = 1;
    let currentPublishAttempt: number = 0;
    let npmDeploymentSucceeded = false;

    while (!npmDeploymentSucceeded && currentPublishAttempt < maxPublishAttempts){
        currentPublishAttempt = currentPublishAttempt + 1;
        console.log(`Publish to NPM Attempt [${currentPublishAttempt}]/[${maxPublishAttempts}]`);

        const nextNpmPackageVersion = getNextNpmPackageVersion.getNextNpmPackageVersion("@microsoft/office-js", packageTag);
        console.log(`Update Package Version: [${nextNpmPackageVersion}]`);
        updatePackageVersion(packageJsonPath, nextNpmPackageVersion);
        
        console.log(`Publish: tag:[${packageTag}]`)
        const tagParameter = packageTag === undefined ? "" : `--tag ${packageTag}`;

        try {
            executeCommand(`npm publish ${tagParameter}`, packageDirectory, false);
            npmDeploymentSucceeded = true;
        } catch (e) {
            const wasFailureDueToPreviouslyPublishedDeletedVersion =
                (e as AdditionalInfoError).additionalInfo &&
                isPublishOverPreviouslyPublishVersionErrorString((e as AdditionalInfoError).additionalInfo);

            if (!wasFailureDueToPreviouslyPublishedDeletedVersion) {
                throw e;
            }
        }
    }
}