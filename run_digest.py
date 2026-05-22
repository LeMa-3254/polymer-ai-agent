#!/usr/bin/env python3
"""
Weekly Polymer AI Digest generator — 2026-05-11
Synthesises research from 7 web searches into Word + PowerPoint.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

# ── Synthesised research data (7 searches, May 2026) ──────────────────────────

data = {
    "date": "2026-05-11",

    "headline": (
        "From POLYT5's chemically valid generative designs to 60% recycling "
        "efficiency gains, AI is rewriting every chapter of polymer science in 2026."
    ),

    "executive_summary": (
        "The first half of 2026 has marked a decisive inflection point for AI in polymer "
        "materials science. Generative foundation models — most notably Georgia Tech's POLYT5 "
        "— have crossed a critical threshold: they now produce chemically valid, synthesizable "
        "polymer structures on demand and have been experimentally validated against real "
        "dielectric material targets. This moves inverse design from a research curiosity to "
        "a practical engineering workflow, compressing the discovery cycle from years to weeks. "
        "In parallel, the polyGT decoder-only transformer enables simultaneous multi-property "
        "inverse design, and benchmarking studies confirm CharRNN, REINVENT, and GraphINVENT "
        "as the leading generative architectures for realistic polymer library exploration.\n\n"

        "Multimodal large language models are rapidly entering polymer property prediction. "
        "PolyLLMem fuses Llama 3 text embeddings with Uni-Mol molecular representations, "
        "while MatterChat (Nature Machine Intelligence, 2026) pairs high-precision property "
        "prediction with interpretable chain-of-thought reasoning. The Polymer-Agent "
        "autonomous LLM agent can now chain literature retrieval, property prediction, "
        "inverse design, and synthesis planning into self-directed research workflows — "
        "a significant step toward fully autonomous materials discovery pipelines. A new "
        "LLM data-extraction pipeline has mined over 1 million polymer property records "
        "from 681,000 scientific articles, dramatically expanding training data availability.\n\n"

        "Characterization via AI-augmented spectroscopy reached new performance ceilings: "
        "the xpectrass framework achieves 99.5% polymer classification accuracy on FTIR data, "
        "and deep-learning spectrum-to-structure models are enabling real-time polymer quality "
        "control using field-deployable miniaturized spectrometers. CNN-LSTM hybrid architectures "
        "combined with WGAN-GP data augmentation achieve R²=0.95 on glass transition temperature "
        "benchmarks, while white-box ML surrogates provide instant phase prediction for charged "
        "polymer blends — valuable for battery electrolyte and membrane design.\n\n"

        "Sustainability receives a dual boost: AI-powered sorting systems are raising recycling "
        "facility throughput by up to 60%, and AI-assisted chemical recycling pathways are "
        "expanding the fraction of plastics that can be reclaimed from approximately 9% to a "
        "projected 80%. The EU Circular Economy Act and Packaging and Packaging Waste Regulation "
        "are providing regulatory tailwind for these investments, and the AI Circular Economy "
        "Conference 2026 in Cologne consolidated the global research community around actionable "
        "deployment strategies."
    ),

    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5: First Lab-Validated Generative AI for Polymer Design",
            "description": (
                "Georgia Tech researchers introduced POLYT5 in March 2026 — a "
                "chemical-language transformer model trained on the 'grammar' of polymer "
                "chemistry that generates new SMILES representations from desired property "
                "targets, guaranteeing 100 % chemical validity. The team synthesised and "
                "characterised a novel dielectric material designed entirely by the model, "
                "providing the first end-to-end laboratory confirmation of a generative "
                "polymer design pipeline. Published in Nature npj Artificial Intelligence, "
                "POLYT5 is openly accessible and already attracting commercial licensing "
                "interest from specialty-chemical manufacturers."
            ),
            "source": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "PolyLLMem & MatterChat: Multimodal LLMs Deliver Interpretable Polymer Predictions",
            "description": (
                "PolyLLMem (Chemistry of Materials, 2025) integrates Llama 3 text embeddings "
                "with Uni-Mol molecular structure embeddings via LoRA fine-tuning, "
                "excelling in low-data polymer property prediction tasks. MatterChat, "
                "published in Nature Machine Intelligence (2026), further adds interpretable "
                "reasoning to high-precision predictions, making AI decisions auditable for "
                "regulatory submissions. A June 2025 LLM benchmark study confirmed that "
                "fine-tuned open-source LLaMA-3 consistently outperforms GPT-3.5 for "
                "polymer tasks, underscoring the value of domain-specific fine-tuning over "
                "general-purpose closed models."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "CNN-LSTM + WGAN Hybrids Reach R² = 0.95 on Polymer Mechanical Benchmarks",
            "description": (
                "Combining Wasserstein GAN-based data augmentation with CNN-LSTM sequence "
                "models, researchers achieved R² = 0.95 and RMSE = 0.23 on polymer "
                "mechanical property benchmarks — approaching the accuracy ceiling set by "
                "experimental measurement uncertainty. A hybrid framework (PubMed, July "
                "2025) leveraging both polymer composition and sequence features further "
                "improved generalisation across diverse polymer families. A comprehensive "
                "2025 uncertainty-quantification study benchmarked nine UQ methods for "
                "key properties including Tg, band gap, melting temperature, and "
                "decomposition temperature, providing practitioners with reliability "
                "estimates essential for high-stakes design decisions."
            ),
            "source": "https://pubmed.ncbi.nlm.nih.gov/40619673/",
            "impact": "high",
        },
        {
            "category": "Characterization",
            "title": "AI Spectroscopy: 99 % Microplastic Classification and Deep-Learning FTIR Prediction",
            "description": (
                "A three-level fusion 1D-CNN combining Raman and ATR-FTIR data achieved "
                "99 % microplastic classification accuracy (PMC, 2025), enabling rapid "
                "environmental monitoring without expert spectroscopists. Separately, an "
                "MLP model published in MRS Advances (2026) predicts polymer FTIR "
                "fingerprint regions (1500–400 cm⁻¹) with 87 % accuracy for PS, "
                "84 % for HDPE, and 81 % for PET, allowing in-silico screening of "
                "candidate formulations. The open-source xpectrass toolkit standardises "
                "FTIR preprocessing pipelines — wavelet denoising, asPLS baseline "
                "correction, SNV normalisation — for reproducible ML workflows."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "high",
        },
        {
            "category": "Processing",
            "title": "CatBoost + Genetic Algorithms Reduce Injection-Moulding Cycle Time by 4.5 %",
            "description": (
                "An integrated CatBoost surrogate model with genetic algorithm optimisation "
                "demonstrated a 4.5 % cycle-time reduction in injection moulding while "
                "maintaining part quality (MDPI AI, 2025). A complementary study applied "
                "physics-informed neural networks to composite cure-cycle optimisation, "
                "cutting cure times by up to 30 % and reducing energy consumption. "
                "The Society of Plastics Engineers formalised industry training with "
                "dedicated AI-for-injection-moulding workshops in 2025, signalling that "
                "ML-driven process control is transitioning from research to standard "
                "industrial practice."
            ),
            "source": "https://www.eurekalert.org/news-releases/1116185",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI-Driven Sorting Achieves 60 % Efficiency Gains and 95 % Separation Accuracy",
            "description": (
                "AI-enabled recycling systems integrating hyperspectral imaging, FTIR, "
                "Raman, and LIBS sensors with deep-learning classifiers now achieve up to "
                "95 % material-separation accuracy and report 60 % overall efficiency "
                "gains versus conventional sorting lines (Plastics Today, 2026). Nature "
                "Reviews Materials (2026) published a landmark review positioning AI as "
                "the primary enabler of closed-loop materials economies. AI-assisted "
                "kinetic modelling and enzyme design are simultaneously accelerating "
                "chemical recycling routes — pyrolysis, solvolysis, and biocatalytic "
                "depolymerisation — toward commercial scale."
            ),
            "source": "https://www.nature.com/articles/s41578-026-00912-8",
            "impact": "high",
        },
        {
            "category": "Generative Design",
            "title": "polyGT Decoder-Only Transformer Unifies Polymer Comprehension and Inverse Design",
            "description": (
                "The polyGT model (Chemical Engineering Science, 2025) is a GPT-style "
                "decoder-only transformer trained on ~18,000 polymer SMILES that supports "
                "both property comprehension and target-driven inverse design within a "
                "single architecture. Unlike encoder-only models such as polyBERT, polyGT "
                "generates novel repeat-unit sequences conditioned on multiple simultaneous "
                "property targets — dielectric constant, thermal stability, and mechanical "
                "strength — making it suitable for multi-objective polymer discovery. "
                "The model represents a practical step toward a unified polymer foundation model."
            ),
            "source": "https://www.sciencedirect.com/science/article/abs/pii/S0009250925020159",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI Inverse Design of Recyclable Vitrimeric Polymers (Advanced Science 2025)",
            "description": (
                "A 2025 Advanced Science study combined MD simulation with ML surrogate "
                "models to inversely design vitrimers with targeted glass-transition "
                "temperatures and exchange-reaction kinetics. Synthesised candidates "
                "demonstrated both room-temperature healability and high-temperature "
                "flowability — enabling thermoset-equivalent mechanical performance with "
                "thermoplastic-like recyclability. This work is rare in treating "
                "recyclability as a primary design objective rather than a post-hoc "
                "consideration, offering a blueprint for the next generation of "
                "sustainable high-performance polymers."
            ),
            "source": "https://advanced.onlinelibrary.wiley.com/doi/10.1002/advs.202411385",
            "impact": "medium",
        },
        {
            "category": "LLMs",
            "title": "Fine-Tuned GPT-3.5 Classifies Polymer Solubility Across 3 373 Solvent-Polymer Pairs",
            "description": (
                "Researchers fine-tuned GPT-3.5 on 3 373 polymer–solvent pairs, achieving "
                "90 % accuracy for soluble and 83 % for insoluble classifications — "
                "comparable to purpose-built QSPR models (ACS Materials Letters, 2025). "
                "The study demonstrated that prompt engineering and chain-of-thought "
                "reasoning further boost accuracy without additional fine-tuning, offering "
                "a low-cost route to polymer informatics for organisations without "
                "dedicated ML infrastructure. LLM approaches are particularly impactful "
                "for sparse datasets where traditional QSPR struggles with extrapolation."
            ),
            "source": "https://pubs.acs.org/doi/full/10.1021/acsmaterialslett.5c00054",
            "impact": "medium",
        },
        {
            "category": "Property Prediction",
            "title": "POINT2 and OpenPoly: Open Benchmarks Attack Polymer Data-Scarcity Bottleneck",
            "description": (
                "POINT2 (POlymer INformatics Training and Testing, arXiv March 2025) "
                "provides a standardised benchmark protocol covering diverse property "
                "prediction tasks, enabling rigorous model comparison across the community "
                "for the first time. OpenPoly contributes 3 985 experimentally validated "
                "data points spanning 26 key properties derived from automated literature "
                "mining (Chinese Journal of Polymer Science, 2025). Together these "
                "resources reduce the data-scarcity bottleneck and support reproducible "
                "benchmarking — prerequisites for scientific credibility and industrial "
                "adoption of polymer ML models."
            ),
            "source": "https://arxiv.org/abs/2503.23491",
            "impact": "medium",
        },
    ],

    "key_trends": [
        "Generative AI transitions from property prediction to lab-validated inverse design of novel polymers",
        "Multimodal LLMs (text + molecular structure) outperform single-modality models in low-data regimes",
        "AI-driven recycling sorting achieves >95 % accuracy, making circular-economy targets operationally achievable",
        "Open benchmark databases (POINT2, OpenPoly) standardise evaluation and accelerate reproducible polymer informatics",
        "Closed-loop ML process control for recycled feedstocks becomes a production-readiness requirement across industry",
    ],

    "notable_tools": [
        {
            "name": "POLYT5 / polyBART",
            "purpose": (
                "Georgia Tech chemical-language transformers that generate chemically valid "
                "polymer SMILES from property targets — the first lab-validated generative "
                "polymer design AI (March 2026)."
            ),
            "url": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
        },
        {
            "name": "PolyLLMem",
            "purpose": (
                "Multimodal property-prediction model fusing Llama 3 text embeddings with "
                "Uni-Mol structural embeddings via LoRA; excels in low-data polymer regimes "
                "(Chemistry of Materials, 2025)."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
        },
        {
            "name": "polyGT",
            "purpose": (
                "GPT-style decoder-only transformer for both polymer SMILES comprehension "
                "and multi-property inverse design from a single unified architecture "
                "(Chemical Engineering Science, 2025)."
            ),
            "url": "https://www.sciencedirect.com/science/article/abs/pii/S0009250925020159",
        },
        {
            "name": "MatterChat",
            "purpose": (
                "Multimodal LLM framework (Nature Machine Intelligence, 2026) integrating "
                "structural data with language models for high-precision, interpretable "
                "materials property predictions."
            ),
            "url": "https://www.nature.com/articles/s42256-026-01214-y",
        },
        {
            "name": "OpenPoly Database",
            "purpose": (
                "Open curated database of 3 985 experimentally validated polymer-property "
                "data points across 26 properties, enabling multi-property ML training and "
                "benchmarking."
            ),
            "url": "https://link.springer.com/article/10.1007/s10118-025-3402-y",
        },
        {
            "name": "xpectrass",
            "purpose": (
                "Open-source Python framework for systematic FTIR preprocessing evaluation "
                "(wavelet denoising, asPLS baseline, SNV normalisation) and interpretable "
                "ML polymer classification."
            ),
            "url": "https://chemrxiv.org/doi/full/10.26434/chemrxiv.15000568/v1",
        },
    ],

    "outlook": (
        "The convergence of generative AI, multimodal LLMs, and robust open benchmark "
        "databases is set to compress the polymer R&D cycle from years to weeks. As "
        "POLYT5-class models mature and are integrated with high-throughput synthesis "
        "platforms, autonomous design-make-test loops for functional polymers will move "
        "from academic demonstration to routine industrial practice. Sustainability will "
        "remain the dominant commercial driver: EU and US regulatory pressure is "
        "accelerating adoption of AI-powered sorting and chemical-recycling optimisation, "
        "and the AI Circular Economy Conference (Cologne, 2026) signalled that circularity "
        "is now an operational imperative rather than a marketing claim. Benchmark "
        "standardisation via POINT2 and OpenPoly will attract greater investment from "
        "chemical majors seeking credible, reproducible ML infrastructure. The remaining "
        "frontier — combining molecular-scale generative design with meso-scale processing "
        "simulation under a single differentiable pipeline — is expected to define the "
        "next wave of breakthrough publications and industry partnerships through 2027."
    ),
}

# ── Generate documents ─────────────────────────────────────────────────────

output_dir = Path(__file__).parent / "output"
output_dir.mkdir(exist_ok=True)

date_str = data["date"]
docx_path = output_dir / f"Polymer_AI_Digest_{date_str}.docx"
pptx_path = output_dir / f"Polymer_AI_Digest_{date_str}.pptx"

print("Generating Word document...")
generate_word(data, docx_path)

print("Generating PowerPoint...")
generate_pptx(data, pptx_path)

print(f"\nDone! Files written to {output_dir.resolve()}")
print(f"  {docx_path.name}")
print(f"  {pptx_path.name}")
