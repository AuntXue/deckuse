/**
 * Move command - set shape position by left/top
 */

import { openWorkspace } from '../core/open-workspace.js'
import { SelectorParser, SelectorResolver } from '../core/selector.js'
import { PptxWriter } from '../writers/pptx-writer.js'
import { commitCommand } from './commit.js'
import { logger } from '../utils/logger.js'
import { CommandError } from '../utils/errors.js'
import { findIdPathInSpTree } from '../utils/pptx-text.js'

export interface MoveOptions {
  left: number
  top: number
  output?: string
}

export async function moveCommand(
  workspaceDir: string,
  selectorStr: string,
  options: MoveOptions
): Promise<void> {
  let opened: Awaited<ReturnType<typeof openWorkspace>> | null = null
  try {
    opened = await openWorkspace(workspaceDir)
    const workspace = opened.workspace
    const writer = new PptxWriter(workspace.workspaceDir)

    const { left, top } = options
    if (!Number.isFinite(left) || left < 0 || left > 1) {
      throw new Error('left must be a float between 0 and 1')
    }
    if (!Number.isFinite(top) || top < 0 || top > 1) {
      throw new Error('top must be a float between 0 and 1')
    }

    const selector = SelectorParser.parse(selectorStr)

    // Fast path: canonical idPath selector like "slide:1/#6,9"
    // SelectorResolver.resolve() currently returns the slide element for slide selectors,
    // so we bypass it here and directly target the given idPath.
    if (selector.type === 'slide' && selector.index !== undefined && selector.idPath?.length) {
      await writer.moveShape(selector.index, `#${selector.idPath.join(',')}`, left, top)
      await writer.save()
      await workspace.updateMetadata({ lastModified: new Date() })

      logger.success(`\nMoved 1 shape(s) to left=${left}, top=${top}`)

      if (opened.mode === 'pptx') {
        if (!options.output) {
          throw new Error('When workspace is a .pptx file, -o/--output is required')
        }
        await commitCommand(workspace.workspaceDir, { output: options.output })
        logger.success(`Output PPTX: ${options.output}`)
      } else if (options.output) {
        await commitCommand(workspace.workspaceDir, { output: options.output })
        logger.success(`Output PPTX: ${options.output}`)
      } else {
        logger.info(`Run 'deckuse commit ${workspaceDir}' to build the PPTX file`)
      }
      return
    }

    const targets = await SelectorResolver.resolve(
      selector,
      workspace.metadata,
      workspace.workspaceDir
    )

    if (targets.length === 0) {
      logger.error(`No matches found for selector: ${selectorStr}`)
      return
    }

    logger.info(`Found ${targets.length} target(s)`)

    let updatedCount = 0
    for (const target of targets) {
      const pageType = target.pageType ?? 'slide'
      if (pageType !== 'slide') {
        logger.warn(`Skipping ${pageType}:${target.slide} (only slide shapes are supported)`)
        continue
      }

      const slide = workspace.getSlide(target.slide)
      if (!slide?.spTree) {
        logger.warn(`Skipping slide ${target.slide} (shape tree not found)`)
        continue
      }

      const idPath = findIdPathInSpTree(slide.spTree, target.element)
      if (!idPath || idPath.length === 0) {
        logger.warn(`Skipping target at slide ${target.slide} (id path not found)`)
        continue
      }

      await writer.moveShape(target.slide, `#${idPath.join(',')}`, left, top)
      updatedCount++
    }

    if (updatedCount === 0) {
      logger.warn('No shapes were moved')
      return
    }

    await writer.save()
    await workspace.updateMetadata({ lastModified: new Date() })

    logger.success(`\nMoved ${updatedCount} shape(s) to left=${left}, top=${top}`)

    if (opened.mode === 'pptx') {
      if (!options.output) {
        throw new Error('When workspace is a .pptx file, -o/--output is required')
      }
      await commitCommand(workspace.workspaceDir, { output: options.output })
      logger.success(`Output PPTX: ${options.output}`)
    } else if (options.output) {
      await commitCommand(workspace.workspaceDir, { output: options.output })
      logger.success(`Output PPTX: ${options.output}`)
    } else {
      logger.info(`Run 'deckuse commit ${workspaceDir}' to build the PPTX file`)
    }
  } catch (error) {
    if (error instanceof Error) {
      throw new CommandError('move', error.message)
    }
    throw error
  } finally {
    if (opened?.mode === 'pptx') {
      await opened.cleanup().catch(() => {})
    }
  }
}
