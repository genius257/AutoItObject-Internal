# Changelog
All notable changes to this project will be documented in this file

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.0] - 2018-07-07
### Added
- This CHANGELOG.md file
- CONTRIBUTING.md file
- Tests
- __get
- __set
- __assign
- __freeze
- __isFrozen
- __isSealed
- __preventExtensions
- __isExtensible
- __lookupGetter
- __lookupSetter
- __seal
- __case
- internal use function __AOI_PropertyGetFromName
- internal use function __AOI_PropertyGetFromId
- internal use function __AOI_PropertyCreate
- internal use function __AOI_GetPtrOffset
- internal use function __AOI_GetPtrValue

### Changed
- Moved project to it's own repository
- __defineGetter and __defineSetter now also supports strings as second argument
- __destructor now also supports strings as first argument
- Split docs from readme into /docs/index.md
- documentation blocks are now following a more standardized docblock standard.

### Removed
- __lock

### Fixed
- desctructor pointer wrong offset after lock propterty was added to the object structure.

## [1.0.3] - 2017-08-11
### Added
- __desctructor
- Garbage collection, with the exception of interface methods

## [1.0.2] - 2017-08-09
### Added
- Parameters for assigning custom methods as IDispatch object internal functionality
- __keys
- PrivateProperty method for ease of use, when making private IDispatch object properties

## [1.0.1] - 2016-08-06
### Added
- Local scope for internal variables used by IDispatch object functionality

## [1.0.0] - 2016-08-06
### Added
- README.md file
- __unset
- __lock
- Accessors can set error code via setError
- Supprt for __all__ AutoIt variable types

## [0.1.2] - 2016-12-21
### Added
- __defineMethod
- Method support for objects

## [0.1.1] - 2016-11-25
### Changed
- Fix script breaking typo in Idispatch function, reported by [Chimp on the AutoIt forum thread](https://www.autoitscript.com/forum/topic/185720-autoitobject-pure-autoit/#elComment_1333941)

## [0.1.0] - 2016-11-24
### Added
- First iteration of the AutoItObject_Internal library
- Accessors via __defineGetter and __defineSetter
