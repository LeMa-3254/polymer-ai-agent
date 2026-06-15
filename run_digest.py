#!/usr/bin/env python3
"""Weekly AI-for-Polymers digest generator — 2026-06-15"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

DATE = "2026-06-15"

data = {
    "date": DATE,
    "headline": (
        "Generative AI clears the lab test for dielectric polymers while autonomous "
        "multi-agent systems and the OPoly26 mega-dataset reshape the polymer "
        "informatics landscape."
    ),
    "executive_summary": (
        "The week of 15 June 2026 consolidates three major advances that signal "
        "polymer AI's transition from bench research to industrial deployment. "
        "Georgia Tech's POLYT5 — the first encoder-decoder foundation model for "
        "generative polymer design to be physically lab-validated — designed a "
        "dielectric polymer that passed experimental specifications on its first "
        "attempt, proving that AI inverse design can replace conventional trial-and-error. "
        "Simultaneously, an arXiv preprint (2602.00103) demonstrated a fully autonomous "
        "multi-agent AI system performing the complete polymer informatics pipeline — "
        "property prediction, generative design, and evaluation — across both synthetic "
        "and bio-polymers without any human-in-the-loop steps.\n\n"
        "The Open Polymers 2026 (OPoly26) dataset, with 6.35 million DFT calculations "
        "spanning over 1.2 billion atoms, provides the community's largest open benchmark "
        "for property-prediction model training and evaluation. Paired with the new TROPIC "
        "database (Faraday Discussions, RSC 2026) — curated thermodynamic data for "
        "ring-opening polymerisation plastics including PLA, PCL, and polycarbonates — "
        "these resources are lowering the barrier to data-driven chemical recycling and "
        "retrosynthesis planning.\n\n"
        "Characterisation automation is maturing rapidly. A 2026 MRS Advances deep "
        "learning model predicts FTIR fingerprint-region spectra for polystyrene (R²=0.87), "
        "HDPE (0.84), and PET (0.81), enabling non-destructive, high-throughput polymer "
        "identification directly applicable to recycling stream sorting. On the processing "
        "side, physics-informed neural networks applied to injection moulding are cutting "
        "cycle times by up to 30% and real-time ML controllers reduce off-spec production "
        "by over 2% while cutting natural gas consumption 10-20%.\n\n"
        "Large language models are crossing from benchmarks into deployment. A multimodal "
        "LLM purpose-built for materials science appeared in Nature Machine Intelligence "
        "in 2026, integrating molecular structure embeddings with natural-language reasoning. "
        "ACS studies confirm that fine-tuned GPT-3.5 matches traditional ML for polymer "
        "solubility prediction even on small datasets. The AI Circular Economy Conference "
        "(Cologne, March 2026) established that AI-driven sorting is delivering up to 60% "
        "efficiency gains in plastic recycling plants, and multi-objective ML optimisation "
        "of pyrolysis is raising oil yields by up to 18%, making chemical recycling "
        "economically viable under 2030 regulatory targets."
    ),
    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5: First Lab-Validated Generative AI Model for Polymer Design",
            "description": (
                "Georgia Tech researchers developed POLYT5, an encoder-decoder foundation "
                "chemical-language model published in npj Artificial Intelligence that "
                "generates polymer structures from target property specifications. A "
                "POLYT5-designed dielectric polymer passed physical laboratory validation "
                "on its first attempt — the first confirmed instance of an AI-generated "
                "polymer meeting experimental targets without iterative human correction. "
                "The system couples AI candidate generation with automated flow chemistry "
                "and real-time analytical feedback, eliminating human bottlenecks in "
                "synthesis cycles."
            ),
            "source": "https://phys.org/news/2026-03-generative-ai-polymer-lab-dielectric.html",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "Autonomous Multi-Agent AI Runs Complete Polymer Informatics Pipeline",
            "description": (
                "An arXiv preprint (2602.00103, Feb 2026) presents a high-throughput "
                "autonomous multi-agent system that orchestrates the entire polymer "
                "informatics workflow — literature mining, property prediction, generative "
                "design, and retrosynthesis — across synthetic and bio-polymers without "
                "human intervention. Specialist sub-agents handle each stage and share "
                "structured outputs, enabling screening throughput infeasible for "
                "single-model approaches. End-to-end design cycles for polyelectrolytes "
                "and biodegradable packaging materials were successfully demonstrated."
            ),
            "source": "https://arxiv.org/pdf/2602.00103",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "OPoly26: 6.35 Million DFT Calculations Form the Largest Open Polymer Benchmark",
            "description": (
                "The Open Polymers 2026 (OPoly26) dataset provides 6.35 million density "
                "functional theory calculations on polymer systems totalling over 1.2 billion "
                "atoms. It captures monomer composition, chain architecture, degree of "
                "polymerisation, and solvation environment variations, giving property-prediction "
                "models a shared, high-quality ground truth. The scale supports "
                "transfer-learning approaches for industrial polymer grades where experimental "
                "data remains scarce, and its open release removes a longstanding reproducibility "
                "bottleneck in polymer ML benchmarking."
            ),
            "source": "https://arxiv.org/pdf/2512.23117",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "Bayesian Network–XGBoost Framework Handles Missing Data for PP Alloy Prediction",
            "description": (
                "Published in Journal of Polymer Science (2026), this two-stage ML framework "
                "uses a Bayesian network for guided feature selection followed by XGBoost "
                "regression to predict impact strength and flexural modulus of polypropylene "
                "alloys. Its key innovation is robustness to missing industrial measurement "
                "data — a pervasive constraint in compounding — outperforming standard "
                "imputation-then-regression pipelines and showing readiness for deployment "
                "on plant-floor datasets where instruments occasionally fail."
            ),
            "source": "https://onlinelibrary.wiley.com/doi/10.1002/pol.20250873",
            "impact": "medium",
        },
        {
            "category": "Characterization",
            "title": "Deep Learning Predicts Polymer FTIR Fingerprint Spectra with R² up to 0.87",
            "description": (
                "A 2026 MRS Advances study applied deep learning to predict FTIR spectra "
                "in the fingerprint region (600-1500 cm⁻¹), achieving R²=0.87 for "
                "polystyrene, 0.84 for HDPE, and 0.81 for PET. The approach eliminates "
                "reliance on manually curated spectral libraries and enables automated, "
                "high-throughput polymer identification directly applicable to recycling "
                "stream QC and microplastics monitoring. The open-source xpectrass "
                "framework provides preprocessing and interpretable ML pipelines for "
                "FTIR classification."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "medium",
        },
        {
            "category": "Processing",
            "title": "Physics-Informed Neural Networks Reduce Injection Moulding Cycle Times by 30%",
            "description": (
                "AI-driven process optimisation systems reviewed in MRS Communications (2026) "
                "deploy physics-informed neural networks (PINNs) to predict optimal "
                "heating/cooling curves for injection moulding, cutting cycle times up to 30%. "
                "ML models trained on plant sensor data detect process drift from feedstock "
                "variability or fouling, reducing off-spec production by over 2% and cutting "
                "natural gas consumption 10-20%. Digital twin integration allows manufacturers "
                "to simulate process changes virtually before physical implementation, "
                "minimising risk during grade transitions."
            ),
            "source": "https://link.springer.com/article/10.1557/s43579-026-00972-5",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "TROPIC Database Enables AI-Guided Chemical Recycling of ROP Polymers",
            "description": (
                "The Thermodynamics of Ring-Opening Polymerisation Informatics Collection "
                "(TROPIC), published in RSC Faraday Discussions (2026), provides curated "
                "thermodynamic datasets for PLA, PCL, polycarbonates, and related ROP "
                "plastics. Structured for direct use in ML models, it enables data-driven "
                "prediction of depolymerisation conditions and retrosynthesis planning for "
                "chemical recycling. The database specifically targets polymers for which "
                "mechanical recycling degrades material quality, making chemical "
                "deconstruction the preferred circular route."
            ),
            "source": "https://pubs.rsc.org/en/content/articlehtml/2026/fd/d5fd00098j",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "AI Sorting and Pyrolysis Optimisation Unlock 60% Recycling Efficiency Gains",
            "description": (
                "Deep learning-powered NIR and hyperspectral sorting systems reported at the "
                "AI Circular Economy Conference (Cologne, March 2026) are achieving up to "
                "60% efficiency gains over manual sorting in plastic recycling facilities. "
                "A complementary Polymers (MDPI, 2026) study demonstrates multi-objective "
                "ML optimisation of plastic pyrolysis — maximising oil yield, minimising "
                "energy, and optimising EROI simultaneously — with gains of up to 18% "
                "oil yield over heuristic baselines. Both technologies are being deployed "
                "as prerequisites for meeting EU 2030 recycled-content mandates."
            ),
            "source": "https://doi.org/10.3390/polym18091062",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "Multimodal LLM for Materials Science Published in Nature Machine Intelligence",
            "description": (
                "A multimodal large language model purpose-built for materials science, "
                "published in Nature Machine Intelligence (2026), integrates molecular "
                "structure embeddings with natural-language reasoning for property prediction, "
                "synthesis planning, and literature extraction. ACS benchmarks confirm that "
                "fine-tuned GPT-3.5 matches or exceeds traditional ML classifiers for polymer "
                "solubility prediction even on small datasets (ACS Materials Letters), "
                "significantly lowering the barrier to LLM adoption in polymer labs without "
                "large proprietary training corpora."
            ),
            "source": "https://www.nature.com/articles/s42256-026-01214-y",
            "impact": "high",
        },
        {
            "category": "Generative Design",
            "title": "PolyTAO + Group SELFIES Achieves 100% Chemically Valid Polymer Generation",
            "description": (
                "A ChemRxiv preprint (2026) integrates the Group SELFIES chemical grammar "
                "with the PolyTAO generative model, achieving 100% chemically valid polymer "
                "structure generation — removing a longstanding validity bottleneck that "
                "required expensive post-hoc filtering. IBM's concurrent patent filing "
                "introduces a continuous expert-guided loop where a classification model "
                "ranks candidates, a generative model creates new structures, and results "
                "feed back into ranking, enabling iterative property-directed exploration "
                "without saturation of the chemical space."
            ),
            "source": "https://chemrxiv.org/engage/chemrxiv/article-details/69143361ef936fb4a2b6f532",
            "impact": "medium",
        },
    ],
    "key_trends": [
        "AI inverse design has crossed from simulation to lab validation: AI-designed polymers now pass physical tests on first attempt",
        "Autonomous multi-agent pipelines are replacing human-in-the-loop polymer informatics, enabling mass-screening throughput",
        "Open mega-datasets (OPoly26, TROPIC) are democratising ML pre-training and consolidating community benchmarking",
        "Physics-informed neural networks and digital twins are delivering measurable manufacturing efficiency gains (10-30%)",
        "LLMs now outperform specialist ML on small polymer datasets, making AI property prediction accessible to data-poor labs",
    ],
    "notable_tools": [
        {
            "name": "POLYT5",
            "purpose": (
                "Encoder-decoder foundation language model for generative polymer inverse design "
                "from target properties; first generatively designed polymer to pass lab validation "
                "(Georgia Tech / npj AI, 2026)."
            ),
            "url": "https://phys.org/news/2026-03-generative-ai-polymer-lab-dielectric.html",
        },
        {
            "name": "OPoly26 Dataset",
            "purpose": (
                "Open benchmark of 6.35 million DFT calculations on polymer systems (>1.2B atoms) "
                "for training and evaluating polymer property prediction models at scale."
            ),
            "url": "https://arxiv.org/pdf/2512.23117",
        },
        {
            "name": "TROPIC Database",
            "purpose": (
                "Thermodynamics of Ring-Opening Polymerisation Informatics Collection — curated "
                "thermodynamic data enabling ML-guided chemical recycling of PLA, PCL, and polycarbonates."
            ),
            "url": "https://pubs.rsc.org/en/content/articlehtml/2026/fd/d5fd00098j",
        },
        {
            "name": "xpectrass",
            "purpose": (
                "Open-source Python framework for FTIR preprocessing evaluation, exploratory "
                "spectral analysis, and interpretable ML classification of polymer types."
            ),
            "url": "https://chemrxiv.org/doi/full/10.26434/chemrxiv.15000568/v1",
        },
        {
            "name": "PolyTAO + Group SELFIES",
            "purpose": (
                "Generative polymer framework combining Group SELFIES grammar with the PolyTAO "
                "model to produce 100% chemically valid polymer structures."
            ),
            "url": "https://chemrxiv.org/engage/chemrxiv/article-details/69143361ef936fb4a2b6f532",
        },
        {
            "name": "Polymer Genome",
            "purpose": (
                "Web-based informatics platform for data-driven polymer property prediction "
                "and materials screening across hundreds of target properties."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.jpcc.8b02913",
        },
    ],
    "outlook": (
        "The next 6-12 months will likely see the first commercially attributed AI-designed "
        "polymer formulation — a direct consequence of POLYT5 and similar foundation models "
        "proving experimental validity beyond simulation. Autonomous multi-agent platforms "
        "will absorb routine polymer informatics tasks (benchmark screening, formulation "
        "optimisation, specification matching), freeing chemists for high-value exploratory "
        "synthesis. The OPoly26 and TROPIC databases signal a broader shift toward FAIR-compliant "
        "polymer data ecosystems that lower entry barriers for smaller companies and academic "
        "groups. On sustainability, AI-optimised pyrolysis and chemical depolymerisation will "
        "become essential tools for meeting 2030 recycled-content mandates in Europe and "
        "North America, with the AI Circular Economy Conference framework providing early "
        "policy scaffolding. Large language models will continue maturing toward integrated "
        "lab assistants capable of linking literature knowledge, experimental data, and "
        "generative design in a single conversational interface — transforming how polymer "
        "R&D teams interact with their own proprietary datasets."
    ),
}

if __name__ == "__main__":
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    docx_path = output_dir / f"Polymer_AI_Digest_{DATE}.docx"
    pptx_path = output_dir / f"Polymer_AI_Digest_{DATE}.pptx"

    print("Generating Word document...")
    generate_word(data, docx_path)

    print("Generating PowerPoint...")
    generate_pptx(data, pptx_path)

    print(f"\nFiles written to {output_dir.resolve()}")
    print(f"  {docx_path.name}")
    print(f"  {pptx_path.name}")
